using System;
using System.Collections;
using System.Linq;
using System.Windows.Forms;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Taxonomy;

namespace MMDNavigator
{
    class MMDHelper
    {

        public static string errMsg = "";


        /// <summary>
        /// </summary>
        public static object GetObj(string siteUrl, int level, TreeNode tNode)
        {
            errMsg = "";

            try
            {
                using (SPSite site = new SPSite(siteUrl))
                {
                    TaxonomySession txsn = new TaxonomySession(site, true);

                    if (level == 0)
                    {
                        // load termstore
                        string termStoreGuid = tNode.Name;

                        TermStore tStore = txsn.TermStores.FirstOrDefault(x => x.Id == new Guid(termStoreGuid));

                        return tStore;

                    }
                    else if (level == 1)
                    {
                        // load termgroup
                        string termStoreGuid = tNode.Parent.Name;
                        string termGroupGuid = tNode.Name;

                        TermStore tStore = txsn.TermStores.FirstOrDefault(x => x.Id == new Guid(termStoreGuid));
                        Group tGroup = tStore.Groups.FirstOrDefault(x => x.Id == new Guid(termGroupGuid));

                        return tGroup;

                    }
                    else if (level == 2)
                    {
                        // load termset
                        string termStoreGuid = tNode.Parent.Parent.Name;
                        string termGroupGuid = tNode.Parent.Name;
                        string termSetGuid = tNode.Name;

                        TermStore tStore = txsn.TermStores.FirstOrDefault(x => x.Id == new Guid(termStoreGuid));
                        Group tGroup = tStore.Groups.FirstOrDefault(x => x.Id == new Guid(termGroupGuid));
                        TermSet tSet = tGroup.TermSets.FirstOrDefault(x => x.Id == new Guid(termSetGuid));

                        return tSet;

                    }
                    else
                    {
                        // load term
                        SortedList slTerms = new SortedList();
                        TreeNode curNode = tNode;

                        while (true)
                        {
                            slTerms.Add(curNode.Level, curNode.Name); // level:guid

                            if (curNode.Level == 0)
                            {
                                break;
                            }
                            else
                            {
                                curNode = curNode.Parent;
                            }
                        }

                        TermStore tStore = txsn.TermStores.FirstOrDefault(x => x.Id == new Guid(slTerms[0].ToString()));
                        Group tGroup = tStore.Groups.FirstOrDefault(x => x.Id == new Guid(slTerms[1].ToString()));
                        TermSet tSet = tGroup.TermSets.FirstOrDefault(x => x.Id == new Guid(slTerms[2].ToString()));
                        Term curTerm = null;

                        int ix = 3;

                        while (ix <= slTerms.Count - 1)
                        {
                            if (ix == 3)
                            {
                                curTerm = tSet.Terms.FirstOrDefault(x => x.Id == new Guid(slTerms[ix].ToString()));
                            }
                            else
                            {
                                curTerm = curTerm.Terms.FirstOrDefault(x => x.Id == new Guid(slTerms[ix].ToString()));
                            }
                            ix++;
                        }

                        return curTerm;

                    }

                }                

            }
            catch(Exception exc)
            {
                errMsg = exc.Message;
            }

            return null;
        }


        /// <summary>
        /// Add nodes to tNode and expand.
        /// </summary>
        public static void LoadChildObjects(string siteUrl, int level, TreeNode tNode)
        {
            errMsg = "";

            try
            {
                using (SPSite site = new SPSite(siteUrl))
                {
                    TaxonomySession txsn = new TaxonomySession(site, true);

                    if (level == 0)
                    {
                        // termstore selected, load groups
                        TermStore tStore = GetObj(siteUrl, level, tNode) as TermStore;

                        if (!string.IsNullOrEmpty(MMDHelper.errMsg))
                        {
                            throw new Exception(MMDHelper.errMsg);
                        }

                        tNode.Nodes.Clear();

                        foreach (Group tGroup in tStore.Groups)
                        {
                            TreeNode tn = new TreeNode();
                            tn.Name = tGroup.Id.ToString();
                            tn.Text = string.Format("{0} [{1}]", tGroup.Name, tGroup.TermSets.Count.ToString());

                            tNode.Nodes.Add(tn);
                        }

                        tNode.Expand();

                    }
                    else if (level == 1)
                    {
                        // termgroup selected, load termsets
                        Group tGroup = GetObj(siteUrl, level, tNode) as Group;

                        if (!string.IsNullOrEmpty(MMDHelper.errMsg))
                        {
                            throw new Exception(MMDHelper.errMsg);
                        }

                        tNode.Nodes.Clear();

                        foreach (TermSet tSet in tGroup.TermSets)
                        {
                            TreeNode tn = new TreeNode();
                            tn.Name = tSet.Id.ToString();
                            tn.Text = string.Format("{0} [{1}]", tSet.Name, tSet.Terms.Count.ToString());

                            tNode.Nodes.Add(tn);
                        }

                        tNode.Expand();

                    }
                    else if (level == 2)
                    {
                        // termset selected, load terms
                        TermSet tSet = GetObj(siteUrl, level, tNode) as TermSet;

                        if (!string.IsNullOrEmpty(MMDHelper.errMsg))
                        {
                            throw new Exception(MMDHelper.errMsg);
                        }

                        //_lstTerm.Clear();
                        tNode.Nodes.Clear();

                        foreach (Term term in tSet.Terms)
                        {
                            TreeNode tn = new TreeNode();
                            tn.Name = term.Id.ToString();
                            tn.Text = string.Format("{0} [{1}]", term.Name, term.Terms.Count.ToString());

                            tNode.Nodes.Add(tn);
                        }

                        tNode.Expand();

                    }
                    else
                    {
                        // term, load its terms (if any)
                        Term term = GetObj(siteUrl, level, tNode) as Term;

                        if (!string.IsNullOrEmpty(MMDHelper.errMsg))
                        {
                            throw new Exception(MMDHelper.errMsg);
                        }

                        if (term.TermsCount > 0 && term.Terms.Any())
                        {
                            tNode.Nodes.Clear();

                            foreach (Term subTerm in term.Terms)
                            {
                                TreeNode tn = new TreeNode();
                                tn.Name = subTerm.Id.ToString();
                                tn.Text = string.Format("{0} [{1}]", subTerm.Name, subTerm.Terms.Count.ToString());

                                tNode.Nodes.Add(tn);
                            }

                            tNode.Expand();
                        }

                    }

                }

            }
            catch (Exception exc)
            {
                errMsg = exc.Message;
            }

        }


    }

}
