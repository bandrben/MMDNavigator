using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Xml.XPath;
using Microsoft.SharePoint.Taxonomy;

namespace MMDNavigator
{
    class XmlImportHelper
    {

    
        /// <summary>
        /// </summary>
        public static string ProcessXml(string xml, Group tGroup, TreeNode curNode)
        {
            var retMsg = "";
            XElement xTree = null;

            try
            {
                // load xml
                xTree = XElement.Load(new StringReader(xml));
            }
            catch(Exception ex)
            {
                return ex.Message;
            }

            foreach (var termSetElement in xTree.XPathSelectElements("/termset"))
            {
                string newTermSetName = GenUtil.MmdNormalize(GenUtil.SafeXmlAttributeToString(termSetElement, "name"));
                Guid newTermSetId = GenUtil.SafeXmlAttributeToGuid(termSetElement, "id");
                bool newTermSetIsAvailForTagging = GenUtil.SafeXmlAttributeToBool(termSetElement, "isavailfortagging");
                string newTermDescr = GenUtil.SafeXmlAttributeToString(termSetElement, "description");
                bool newTermSetIsOpenForTermCreation = GenUtil.SafeXmlAttributeToBool(termSetElement, "isopenfortermcreation");

                if (GenUtil.IsNull(newTermSetName))
                {
                    return "TermSet name is missing.";
                }

                // create termset (or update if found)
                TermSet tSet = null;
                try
                {
                    tSet = tGroup.TermStore.GetTermSet(newTermSetId);

                    if (tSet != null)
                    {
                        // termset found using guid
                        tSet.Name = newTermSetName;
                        tSet.Description = newTermDescr;
                        tSet.IsAvailableForTagging = newTermSetIsAvailForTagging;
                        tSet.IsOpenForTermCreation = newTermSetIsOpenForTermCreation;
                        tSet.TermStore.CommitAll();

                    }
                    else
                    {
                        tSet = tGroup.TermStore.GetTermSets(newTermSetName, CultureInfo.CurrentCulture.LCID).FirstOrDefault();

                        if (tSet != null)
                        {
                            // termset found using name
                            tSet.Description = newTermDescr;
                            tSet.IsAvailableForTagging = newTermSetIsAvailForTagging;
                            tSet.IsOpenForTermCreation = newTermSetIsOpenForTermCreation;
                            tSet.TermStore.CommitAll();
                        }
                        else
                        {
                            tSet = tGroup.CreateTermSet(newTermSetName, newTermSetId, CultureInfo.CurrentCulture.LCID);
                            tSet.Description = newTermDescr;
                            tSet.IsAvailableForTagging = newTermSetIsAvailForTagging;
                            tSet.IsOpenForTermCreation = newTermSetIsOpenForTermCreation;
                            tSet.TermStore.CommitAll();
                        }
                    }

                }
                catch (Exception ex)
                {
                    return ex.Message;
                }

                // create terms within (recursive)
                try
                {
                    foreach (var termElement in termSetElement.XPathSelectElements("term"))
                    {
                        ProcessTerm(termElement, tSet, null);
                    }
                }
                catch (Exception ex)
                {
                    return ex.Message;
                }
            }

            return retMsg;
        }


        /// <summary>
        /// </summary>
        private static void ProcessTerm(XElement termElement, TermSet tSet, Term term)
        {
            string newTermName = GenUtil.MmdNormalize(GenUtil.SafeXmlAttributeToString(termElement, "name"));
            Guid? newTermId = GenUtil.SafeXmlAttributeToGuidOrNull(termElement, "id");
            bool newTermIsAvailForTagging = GenUtil.SafeXmlAttributeToBool(termElement, "isavailfortagging");
            string newTermDescr = GenUtil.SafeXmlAttributeToString(termElement, "description");
            bool newTermReuse = GenUtil.SafeXmlAttributeToBool(termElement, "reuse");
            bool newTermReuseBranch = GenUtil.SafeXmlAttributeToBool(termElement, "reusebranch");

            if (GenUtil.IsNull(newTermName))
            {
                throw new Exception("Term name is empty.");
            }

            // create term (or get existing term to update, or reuse term)
            Term newTerm = null;
            bool termExists = true;
            bool termIsReused = false;

            if (tSet != null)
            {
                // termset passed to function, the term being worked on is a level 0 term in a termset
                if (newTermReuse && newTermId != null)
                {
                    // try to reuse term using termguid
                    newTerm = tSet.TermStore.GetTerm((Guid)newTermId);

                    if (newTerm != null)
                    {
                        // resuse term
                        newTerm = tSet.ReuseTerm(newTerm, newTermReuseBranch);
                        termIsReused = true;
                        newTerm.TermStore.CommitAll();
                    }
                }
                
                if (!termIsReused)
                {
                    if (newTermId != null)
                    {
                        // try to get term based on guid
                        newTerm = tSet.TermStore.GetTerm((Guid)newTermId);
                    }

                    if (newTermId == null)
                    {
                        // try to get term based on name
                        try
                        {
                            newTerm = tSet.Terms[newTermName];
                        }
                        catch
                        {
                            newTerm = null;
                        }
                    }

                    if (newTerm == null)
                    {
                        // create new term
                        newTerm = tSet.CreateTerm(newTermName, CultureInfo.CurrentCulture.LCID, (newTermId == null ? Guid.NewGuid() : (Guid)newTermId));
                        termExists = false;
                    }
                }
            }
            else
            {
                // termset not passed to function, term being worked on is a level n term in a termset (term within term)
                if (newTermReuse && newTermId != null)
                {
                    // try to reuse term using termguid
                    newTerm = term.TermStore.GetTerm((Guid)newTermId);

                    if (newTerm != null)
                    {
                        // resuse term
                        newTerm = term.ReuseTerm(newTerm, newTermReuseBranch);
                        termIsReused = true;
                        newTerm.TermStore.CommitAll();
                    }
                }

                if (!termIsReused)
                {
                    if (newTermId != null)
                    {
                        // try to get term based on guid
                        newTerm = term.TermSet.GetTerm((Guid)newTermId);
                    }

                    if (newTermId == null)
                    {
                        try
                        {
                            newTerm = term.Terms[newTermName];
                        }
                        catch
                        {
                            newTerm = null;
                        }

                        // try to get term based on name
                        //foreach (var termFound in term.TermSet.GetTerms(newTermName, false))
                        //{
                        //    if (termFound.Parent.Name.Trim().ToLower() == termElement.Parent.Attribute("name").Value.Trim().ToLower())
                        //    {
                        //        newTerm = termFound;
                        //    }
                        //}

                        //newTerm = term.TermSet.GetTerms(newTermName, false).FirstOrDefault();

                        //newTerm = term.GetTerms(newTermName, CultureInfo.CurrentCulture.LCID, true, StringMatchOption.ExactMatch, 1, false).FirstOrDefault();
                    }

                    if (newTerm == null)
                    {
                        // create new term
                        newTerm = term.CreateTerm(newTermName, CultureInfo.CurrentCulture.LCID, (newTermId == null ? Guid.NewGuid() : (Guid)newTermId));
                        termExists = false;
                    }
                }
            }


            // update term properties (not if being reused)
            if (!termIsReused)
            {
                // term not reused
                if (newTerm == null)
                {
                    throw new Exception("Term not found.");
                }

                newTerm.IsAvailableForTagging = newTermIsAvailForTagging;

                if (!GenUtil.IsNull(newTermDescr))
                    newTerm.SetDescription(newTermDescr, CultureInfo.CurrentCulture.LCID);

                // reset labels/synonyms
                if (termExists)
                {
                    int i = 0;
                    while (i < newTerm.Labels.Count)
                    {
                        if (!newTerm.Labels[i].IsDefaultForLanguage)
                        {
                            newTerm.Labels[i].Delete();
                        }
                        else
                        {
                            i++;
                        }
                    }
                }

                // recreate term labels
                foreach (var termLabel in termElement.XPathSelectElements("label"))
                {
                    var lbl = GenUtil.MmdNormalize(GenUtil.SafeXmlAttributeToString(termLabel, "name"));

                    if (!GenUtil.IsNull(lbl) && lbl != newTermName)
                    {
                        newTerm.CreateLabel(lbl, CultureInfo.CurrentCulture.LCID, false);
                    }
                }

                newTerm.TermStore.CommitAll();

            }

            if (termIsReused && newTermReuseBranch)
            {
                // quit if term is reused and using existing term branch
                return;
            }

            // continue processing subterms
            foreach (var subTermElement in termElement.XPathSelectElements("term"))
            {
                ProcessTerm(subTermElement, null, newTerm);
            }

        }

    }
}
