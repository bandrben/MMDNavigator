using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;
using Microsoft.SharePoint.Taxonomy;

namespace MMDNavigator
{
    class XmlExportHelper
    {

        /// <summary>
        /// </summary>
        public static bool ExportToXml(SaveFileDialog saveFileDialog, string siteUrl, TreeNode tNode, out string msg)
        {
            msg = "OK";

            if (tNode == null || tNode.Level < 1 || tNode.Level > 2)
            {
                msg = "Can only export Groups and TermSets.";
                return false;
            }

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                var lstTermSets = new List<TermSet>();

                if (tNode.Level == 1)
                {
                    // export all termsets in termgroup
                    var tGroup = MMDHelper.GetObj(siteUrl, tNode.Level, tNode) as Group;

                    if (!string.IsNullOrEmpty(MMDHelper.errMsg))
                    {
                        msg = MMDHelper.errMsg;
                        return false;
                    }

                    if (tGroup == null)
                    {
                        msg = "Group not found.";
                        return false;
                    }

                    lstTermSets.AddRange(tGroup.TermSets);

                }
                else if (tNode.Level == 2)
                {
                    // export specific termset
                    var termSet = MMDHelper.GetObj(siteUrl, tNode.Level, tNode) as TermSet;

                    if (!string.IsNullOrEmpty(MMDHelper.errMsg))
                    {
                        msg = MMDHelper.errMsg;
                        return false;
                    }

                    if (termSet == null)
                    {
                        msg = "TermSet not found.";
                        return false;
                    }

                    lstTermSets.Add(termSet);

                }

                if (lstTermSets.Count <= 0)
                {
                    msg = "No termsets found to export.";
                    return false;
                }

                var settings = new XmlWriterSettings();
                settings.Indent = true;

                var sb = new StringBuilder();
                var xmlWriter = XmlWriter.Create(sb, settings);

                xmlWriter.WriteStartDocument();
                xmlWriter.WriteStartElement("termsets");

                foreach (var curTermSet in lstTermSets)
                {
                    ExportToXml(ref xmlWriter, curTermSet, null);
                }

                xmlWriter.WriteEndElement();
                xmlWriter.WriteEndDocument();
                xmlWriter.Close();

                var fs = new FileStream(saveFileDialog.FileName, FileMode.Create);
                var writer = new StreamWriter(fs);
                writer.Write(sb.ToString());
                writer.Close();
                fs.Close();

            }

            return msg == "OK";
        }

        /// <summary>
        /// </summary>
        private static void ExportToXml(ref XmlWriter xmlWriter, TermSet termSet, Term term)
        {
            if (termSet != null)
            {
                // export termset
                xmlWriter.WriteStartElement("termset"); // <termset>
                xmlWriter.WriteAttributeString("name", GenUtil.MmdDenormalize(termSet.Name));
                xmlWriter.WriteAttributeString("id", termSet.Id.ToString());
                xmlWriter.WriteAttributeString("description", termSet.Description);
                xmlWriter.WriteAttributeString("isavailfortagging", (termSet.IsAvailableForTagging ? "true" : "false"));
                xmlWriter.WriteAttributeString("isopenfortermcreation", (termSet.IsOpenForTermCreation ? "true" : "false"));

                foreach (var curTerm in termSet.Terms)
                {
                    ExportToXml(ref xmlWriter, null, curTerm);
                }

                xmlWriter.WriteEndElement(); // </termset>

            }
            else if (term != null)
            {
                // export term
                xmlWriter.WriteStartElement("term"); // <term>
                xmlWriter.WriteAttributeString("name", GenUtil.MmdDenormalize(term.Name));
                xmlWriter.WriteAttributeString("id", term.Id.ToString());
                xmlWriter.WriteAttributeString("description", term.GetDescription());
                xmlWriter.WriteAttributeString("isavailfortagging", term.IsAvailableForTagging ? "true" : "false");
                xmlWriter.WriteAttributeString("reuse", !term.IsRoot && term.IsReused ? "true" : "false");
                xmlWriter.WriteAttributeString("reusebranch", "false"); // #todo# how to determine if reusebranch is true?

                if (term.Labels.Count > 0)
                {
                    foreach (var label in term.Labels)
                    {
                        if (label.Value != term.Name)
                        {
                            xmlWriter.WriteStartElement("label"); // <label>
                            xmlWriter.WriteAttributeString("name", GenUtil.MmdDenormalize(label.Value));
                            xmlWriter.WriteEndElement(); // </label>
                        }
                    }
                }

                foreach (var curTerm in term.Terms)
                {
                    ExportToXml(ref xmlWriter, null, curTerm);
                }

                xmlWriter.WriteEndElement(); // </term>

            }
            
        }

    }
}
