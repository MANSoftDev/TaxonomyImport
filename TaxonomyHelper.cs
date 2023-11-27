using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Taxonomy;
using System.Xml.Linq;
using Microsoft.SharePoint;
using System.IO;
using System.Reflection;

namespace MANSoftDev.SharePoint.Utilities.Taxonomy
{
    /// <summary>
    /// Provide helper methods for working with Managed Metadata
    /// </summary>
    public class TaxonomyHelper
    {
        private static string m_TaxonomyRoot;

        /// <summary>
        /// Imports TermSets and Terms from XML
        /// </summary>
        /// <param name="site">SPSite to use for importing</param>
        /// <param name="metadata">TextReader containing taxonomy to import</param>
        public static void ImportTerms(SPSite site, TextReader metadata)
        {
            Site = site;
            XElement taxonomy = XElement.Load(metadata);
            if(taxonomy.Attribute("lcid") != null)
            {
                LCID = Convert.ToInt32(taxonomy.Attribute("lcid").Value);
            }
            else
            {
                LCID = 1033;
            }

            // Get the TaxonomySession from the specified site
            TaxonomySession session = new TaxonomySession(site);

            foreach(XElement termStoreElement in taxonomy.Elements())
            {
                // Get the TermStore from the import file
                TermStore termStore = session.TermStores[termStoreElement.Attribute("name").Value];
                if(termStore != null)
                {
                    foreach(XElement groupElement in termStoreElement.Elements())
                    {
                        // Check if the Group already exits and create it if not
                        Group group = termStore.Groups.FirstOrDefault(e => e.Name == groupElement.Attribute("name").Value);
                        if(group == null)
                        {
                            group = termStore.CreateGroup(groupElement.Attribute("name").Value);
                        }

                        foreach(XElement termsetElement in groupElement.Elements())
                        {
                            CreateTermsSets(termsetElement, group);
                        }
                    }
                }

                // Commit everything that has been added
                termStore.CommitAll();
            }
        }

        #region Private Methods

        /// <summary>
        /// Create TermSets from the given XElement
        /// </summary>
        /// <param name="element">XElement containing TermSets</param>
        /// <param name="group">Group to assign TermSets to</param>
        private static void CreateTermsSets(XElement element, Group group)
        {
            TermSet set = group.TermSets.FirstOrDefault(e => e.Name == element.Attribute("name").Value);
            if(set == null)
            {
                set = group.CreateTermSet(element.Attribute("name").Value);
            }

            foreach(XElement terms in element.Elements())
            {
                CreateTerms(terms, set);
            }
        }

        /// <summary>
        /// Create Terms from the given XElement
        /// </summary>
        /// <param name="element">XElement containing Terms</param>
        /// <param name="set">TermSet to assign Terms to</param>
        private static void CreateTerms(XElement element, TermSet set)
        {
            Term term = set.Terms.FirstOrDefault(e => e.Name == element.Attribute("name").Value);
            if(term == null)
            {
                term = set.CreateTerm(element.Attribute("name").Value, LCID,
                    new Guid(element.Attribute("guid").Value));

                int[] ids = TaxonomyField.GetWssIdsOfTerm(Site, set.Group.TermStore.Id, set.Id, term.Id, false, 999);
                if(ids.Length == 0)
                {
                    int id = CreateWssId(Site, term);
                }
            }
        }

        /// <summary>
        /// Force creation of WssId for given Term
        /// </summary>
        /// <param name="site">Site to create WssId reference</param>
        /// <param name="term">Term to create WssId for</param>
        /// <returns>WssId or -1 if not created</returns>
        private static int CreateWssId(SPSite site, Term term)
        {
            site.RootWeb.AllowUnsafeUpdates = true;

            int result = -1;

            MethodInfo mi = typeof(TaxonomyField).GetMethod("AddTaxonomyGuidToWss",
                    BindingFlags.NonPublic | BindingFlags.Static, null,
                    new Type[3] { typeof(SPSite), typeof(Term), typeof(bool) },
                    null);
            if(mi != null)
            {
                result = (int)mi.Invoke(null, new object[3] { site, term, false });
            }

            site.RootWeb.AllowUnsafeUpdates = false;

            return result;
        }

        #endregion

        #region Properties

        public static string TaxonomyRoot
        {
            get
            {
                if(string.IsNullOrEmpty(m_TaxonomyRoot))
                {
                    m_TaxonomyRoot = "MANSoftDev";
                }
                return m_TaxonomyRoot;
            }
            set { m_TaxonomyRoot = value; }
        }

        private static SPSite Site { get; set; }
        private static int LCID { get; set; }

        #endregion
    }
}
