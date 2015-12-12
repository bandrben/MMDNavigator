using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Taxonomy;

namespace MMDNavigator
{

    public class MsExportObject
    {
        public string TermSetName;
        public string TermSetDescription;
        public string Lcid; // 1033/[null]
        public string AvailableforTagging; // TRUE/FALSE
        public string TermDescription;
        public string Level1Term;
        public string Level2Term;
        public string Level3Term;
        public string Level4Term;
        public string Level5Term;
        public string Level6Term;
        public string Level7Term;
    }

    public class MicrosoftExporter
    {

        /// <summary>
        /// </summary>
        public static bool ExportToMsFormat(SaveFileDialog saveFileDialog, string siteUrl, TreeNode tNode, out string msg)
        {
            msg = "OK";

            try
            {
                if (tNode == null || tNode.Level != 2)
                {
                    msg = "Cannot export, please select a termset";
                    return false;
                }

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    TermSet tSet = MMDHelper.GetObj(siteUrl, tNode.Level, tNode) as TermSet;

                    if (!string.IsNullOrEmpty(MMDHelper.errMsg))
                    {
                        throw new Exception(MMDHelper.errMsg);
                    }

                    List<MsExportObject> lstExObjs = new List<MsExportObject>();

                    int i = 0;
                    foreach (Term term in tSet.Terms)
                    {
                        lstExObjs.Add(new MsExportObject()
                        {
                            TermSetName = (i == 0 ? string.Format("\"{0}\"", tSet.Name) : ""),
                            TermSetDescription = (i == 0 ? (string.IsNullOrEmpty(tSet.Description) ? "" : string.Format("\"{0}\"", tSet.Description)) : ""),
                            Lcid = "",
                            AvailableforTagging = (term.IsAvailableForTagging ? "TRUE" : "FALSE"),
                            TermDescription = (string.IsNullOrEmpty(term.GetDescription()) ? "" : string.Format("\"{0}\"", term.GetDescription())),
                            Level1Term = string.Format("\"{0}\"", term.Name)
                        });

                        if (term.TermsCount > 0)
                        {
                            LoadTerms(lstExObjs, term, 2);
                        }

                        i++;
                    }

                    ExportToCsv(lstExObjs, saveFileDialog.FileName);

                }

            }
            catch (Exception exc)
            {
                msg = exc.Message;
            }

            return (msg == "OK");
        }


        /// <summary>
        /// Recursive function for loading terms.
        /// </summary>
        private static void LoadTerms(List<MsExportObject> lstExObjs, Term term, int level)
        {
            if (level > 7)
            {
                return;
            }

            foreach (Term curTerm in term.Terms)
            {
                // add and recurse
                MsExportObject msExportObject = new MsExportObject();

                msExportObject.AvailableforTagging = (curTerm.IsAvailableForTagging ? "TRUE" : "FALSE");
                msExportObject.TermDescription = (string.IsNullOrEmpty(curTerm.GetDescription())
                                                      ? ""
                                                      : string.Format("\"{0}\"", curTerm.GetDescription()));

                if (level == 2)
                {
                    msExportObject.Level1Term = lstExObjs[lstExObjs.Count - 1].Level1Term;
                    msExportObject.Level2Term = string.Format("\"{0}\"", curTerm.Name);
                }
                else if (level == 3)
                {
                    msExportObject.Level1Term = lstExObjs[lstExObjs.Count - 1].Level1Term;
                    msExportObject.Level2Term = lstExObjs[lstExObjs.Count - 1].Level2Term;
                    msExportObject.Level3Term = string.Format("\"{0}\"", curTerm.Name);
                }
                else if (level == 4)
                {
                    msExportObject.Level1Term = lstExObjs[lstExObjs.Count - 1].Level1Term;
                    msExportObject.Level2Term = lstExObjs[lstExObjs.Count - 1].Level2Term;
                    msExportObject.Level3Term = lstExObjs[lstExObjs.Count - 1].Level3Term;
                    msExportObject.Level4Term = string.Format("\"{0}\"", curTerm.Name);
                }
                else if (level == 5)
                {
                    msExportObject.Level1Term = lstExObjs[lstExObjs.Count - 1].Level1Term;
                    msExportObject.Level2Term = lstExObjs[lstExObjs.Count - 1].Level2Term;
                    msExportObject.Level3Term = lstExObjs[lstExObjs.Count - 1].Level3Term;
                    msExportObject.Level4Term = lstExObjs[lstExObjs.Count - 1].Level4Term;
                    msExportObject.Level5Term = string.Format("\"{0}\"", curTerm.Name);
                }
                else if (level == 6)
                {
                    msExportObject.Level1Term = lstExObjs[lstExObjs.Count - 1].Level1Term;
                    msExportObject.Level2Term = lstExObjs[lstExObjs.Count - 1].Level2Term;
                    msExportObject.Level3Term = lstExObjs[lstExObjs.Count - 1].Level3Term;
                    msExportObject.Level4Term = lstExObjs[lstExObjs.Count - 1].Level4Term;
                    msExportObject.Level5Term = lstExObjs[lstExObjs.Count - 1].Level5Term;
                    msExportObject.Level6Term = string.Format("\"{0}\"", curTerm.Name);
                }
                else if (level == 7)
                {
                    msExportObject.Level1Term = lstExObjs[lstExObjs.Count - 1].Level1Term;
                    msExportObject.Level2Term = lstExObjs[lstExObjs.Count - 1].Level2Term;
                    msExportObject.Level3Term = lstExObjs[lstExObjs.Count - 1].Level3Term;
                    msExportObject.Level4Term = lstExObjs[lstExObjs.Count - 1].Level4Term;
                    msExportObject.Level5Term = lstExObjs[lstExObjs.Count - 1].Level5Term;
                    msExportObject.Level6Term = lstExObjs[lstExObjs.Count - 1].Level6Term;
                    msExportObject.Level7Term = string.Format("\"{0}\"", curTerm.Name);
                }

                lstExObjs.Add(msExportObject);

                if (curTerm.TermsCount > 0)
                {
                    LoadTerms(lstExObjs, curTerm, level + 1);
                }

            }

        }


        /// <summary>
        /// </summary>
        private static void ExportToCsv(List<MsExportObject> lstExObjs, string fileName)
        {
            StringBuilder sb = new StringBuilder("");

            sb.AppendLine(
                "\"Term Set Name\",\"Term Set Description\",\"LCID\",\"Available for Tagging\",\"Term Description\",\"Level 1 Term\",\"Level 2 Term\",\"Level 3 Term\",\"Level 4 Term\",\"Level 5 Term\",\"Level 6 Term\",\"Level 7 Term\"");

            foreach (var msExportObject in lstExObjs)
            {
                sb.AppendLine((string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11}",
                    GenUtil.MmdDenormalize(msExportObject.TermSetName),
                    msExportObject.TermSetDescription,
                    msExportObject.Lcid,
                    msExportObject.AvailableforTagging,
                    msExportObject.TermDescription,
                    GenUtil.MmdDenormalize(msExportObject.Level1Term),
                    GenUtil.MmdDenormalize(msExportObject.Level2Term),
                    GenUtil.MmdDenormalize(msExportObject.Level3Term),
                    GenUtil.MmdDenormalize(msExportObject.Level4Term),
                    GenUtil.MmdDenormalize(msExportObject.Level5Term),
                    GenUtil.MmdDenormalize(msExportObject.Level6Term),
                    GenUtil.MmdDenormalize(msExportObject.Level7Term)
                    )));
            }

            FileStream fs = new FileStream(fileName, FileMode.Create);
            StreamWriter writer = new StreamWriter(fs);
            writer.Write(sb.ToString());
            writer.Close();
            fs.Close();
            
        }

    }
}
