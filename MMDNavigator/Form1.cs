using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Taxonomy;
using System.Collections;
using System.IO;
using System.Globalization;
using System.Configuration;

namespace MMDNavigator
{
    public partial class Form1 : Form
    {

        private const string NewTermGuidLabel = "Enter new GUID (optional)";


        /// <summary>
        /// </summary>
        public Form1()
        {
            InitializeComponent();

            tvMMD.NodeMouseDoubleClick += new TreeNodeMouseClickEventHandler(tvMMD_NodeMouseDoubleClick);
            tvMMD.NodeMouseClick += new TreeNodeMouseClickEventHandler(tvMMD_NodeMouseClick);

            tvMMD.MouseHover += new EventHandler(tvMMD_MouseHover);
            txtSiteUrl.MouseHover += new EventHandler(txtSiteUrl_MouseHover);
            btnExport.MouseHover += new EventHandler(btnExport_MouseHover);
            chkSplitSyns.MouseHover += new EventHandler(chkSplitSyns_MouseHover);
            txtTermLabels.MouseHover += new EventHandler(txtTermLabels_MouseHover);

            rbExportUseMS.MouseHover += new EventHandler(rbExportUseMS_MouseHover);
            rbExportUseXml.MouseHover += new EventHandler(rbExportUseXml_MouseHover);
            rbExportUseProp.MouseHover += new EventHandler(rbExportUseProp_MouseHover);

            tbBulkMergeTerms.MouseHover += tbBulkMergeTerms_MouseHover;

            saveFileDialog1.DefaultExt = "csv";
            saveFileDialog1.Filter = @"csv files (*.csv)|*.csv";
            saveFileDialog1.Title = @"Export in CSV format";

            saveFileDialog2.DefaultExt = "xml";
            saveFileDialog2.Filter = @"xml files (*.xml)|*.xml";
            saveFileDialog2.Title = @"Export in XML format";

            picLogo.Click += new EventHandler(picLogo_Click);
            picLogoWait.Click += new EventHandler(picLogoWait_Click);

            picLogoWait.Visible = false;

            txtNewTermGuid.Text = NewTermGuidLabel;

            txtSiteUrl.Text = GenUtil.SafeTrim(ConfigurationManager.AppSettings["siteUrl"]);

            rbExportUseMS.Click += new EventHandler(rbExportUseMS_Click);
            rbExportUseProp.Click += new EventHandler(rbExportUseProp_Click);
            rbExportUseXml.Click += new EventHandler(rbExportUseXml_Click);

        }


        void rbExportUseXml_Click(object sender, EventArgs e)
        {
            bool state = rbExportUseXml.Checked;

            rbExportUseProp.Checked = !state;
            rbExportUseXml.Checked = state;
            rbExportUseMS.Checked = !state;
        }


        void rbExportUseProp_Click(object sender, EventArgs e)
        {
            bool state = rbExportUseProp.Checked;

            rbExportUseProp.Checked = state;
            rbExportUseXml.Checked = !state;
            rbExportUseMS.Checked = !state;
        }


        void rbExportUseMS_Click(object sender, EventArgs e)
        {
            bool state = rbExportUseMS.Checked;

            rbExportUseProp.Checked = !state;
            rbExportUseXml.Checked = !state;
            rbExportUseMS.Checked = state;
        }


        /// <summary>
        /// </summary>
        void picLogo_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("http://www.bandrsolutions.com/?utm_source=SPCAMLQueryHelper&utm_medium=application&utm_campaign=SPCAMLQueryHelper");
        }


        /// <summary>
        /// </summary>
        private void picLogoWait_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("http://www.bandrsolutions.com/?utm_source=SPCAMLQueryHelper&utm_medium=application&utm_campaign=SPCAMLQueryHelper");
        }


        #region "Hover Tips"


        void rbExportUseProp_MouseHover(object sender, EventArgs e)
        {
            txtHover.Text = string.Format("Export to custom CSV format, choose TermStore, Group, TermSet, Term, or sub-Term in tree.");
        }

        void rbExportUseXml_MouseHover(object sender, EventArgs e)
        {
            txtHover.Text = string.Format("Export to custom XML format, choose Group or TermSet in tree.");
        }

        void rbExportUseMS_MouseHover(object sender, EventArgs e)
        {
            txtHover.Text = string.Format("Export to Microsoft CSV format, choose TermSet in tree.");
        }


        void tbBulkMergeTerms_MouseHover(object sender, EventArgs e)
        {
            txtHover.Text = string.Format("Bulk upload terms into a termset, skips terms if they already exist, extra labels not supported, flat heirarchy only. Term separator is newline.");
        }


        /// <summary>
        /// </summary>
        void txtTermLabels_MouseHover(object sender, EventArgs e)
        {
            txtHover.Text = string.Format("Enter term synonyms, must be different than term name, separate each with semi-colon.");
        }


        /// <summary>
        /// </summary>
        void chkSplitSyns_MouseHover(object sender, EventArgs e)
        {
            txtHover.Text = string.Format("When exporting to CSV, either split each label to a different row, or include additional columns for each label found.");
        }


        /// <summary>
        /// </summary>
        void btnExport_MouseHover(object sender, EventArgs e)
        {
            txtHover.Text = string.Format("Select node to set root of export, lowest level root is TermSet. MS Format only exports 7 levels of terms max.");
        }


        /// <summary>
        /// </summary>
        void txtSiteUrl_MouseHover(object sender, EventArgs e)
        {
            txtHover.Text = string.Format("Enter a site url that uses a Managed Metadata Service Application.");
        }


        /// <summary>
        /// </summary>
        void tvMMD_MouseHover(object sender, EventArgs e)
        {
            txtHover.Text = string.Format("Double-click a node to expand (if available). Single-click a node to load its detail.");
        }


        #endregion


        /// <summary>
        /// Load termstores into tree
        /// </summary>
        private void btnLoadSite_Click(object sender, EventArgs e)
        {
            StartWait();

            var th = new Thread(new ThreadStart(btnLoadSite_Click_Worker));
            th.Start();
        }

        private void btnLoadSite_Click_Worker()
        {
            // do work that does not affect ui here
            var lst = new List<TreeNode>();
            var msg = "";

            try
            {
                using (var site = new SPSite(txtSiteUrl.Text))
                {
                    var txsn = new TaxonomySession(site, true);

                    foreach (TermStore _termStore in txsn.TermStores)
                    {
                        TreeNode tn = new TreeNode();
                        tn.Name = _termStore.Id.ToString();
                        tn.Text = _termStore.Name;

                        lst.Add(tn);
                    }
                }

            }
            catch (Exception exc)
            {
                msg = exc.ToString();
            }

            // update parent thread gui
            this.Invoke((MethodInvoker)delegate
            {
                tvMMD.Nodes.Clear();

                if (msg == "")
                {
                    foreach (TreeNode treeNode in lst)
                    {
                        tvMMD.Nodes.Add(treeNode);
                    }
                    cout("TermStore(s) Loaded.");
                   
                }
                else
                {
                    cout("ERROR", msg);
                }

                StopWait();

            });
        }


        /// <summary>
        /// Expand current node, load sub nodes (groups, termsets, terms)
        /// </summary>
        void tvMMD_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            StartWait();

            this.Invoke((MethodInvoker)delegate
            {

                MMDHelper.LoadChildObjects(txtSiteUrl.Text, e.Node.Level, e.Node);

                if (!string.IsNullOrEmpty(MMDHelper.errMsg))
                {
                    cout("ERROR", MMDHelper.errMsg);
                }

            });

            StopWait();
        }


        /// <summary>
        /// Load Node details (termstore, group, termset, term)
        /// </summary>
        void tvMMD_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            StartWait();

            this.Invoke((MethodInvoker)delegate
            {

                while(true)
                {

                    if (e.Node.Level == 0)
                    {
                        // termstore, load current node detail
                        TermStore tStore = MMDHelper.GetObj(txtSiteUrl.Text, e.Node.Level, e.Node) as TermStore;

                        if (!string.IsNullOrEmpty(MMDHelper.errMsg))
                        {
                            cout("ERROR", MMDHelper.errMsg);
                            break;
                        }

                        txtTermStoreId.Text = tStore.Id.ToString();
                        txtTermStoreName.Text = tStore.Name ?? "";

                        txtCurSelNode.Text = tStore.Name;

                        // select tab
                        tabControl1.SelectTab(tabTermStore);

                    }
                    else if (e.Node.Level == 1)
                    {
                        // termgroup, load current node detail
                        Group tGroup = MMDHelper.GetObj(txtSiteUrl.Text, e.Node.Level, e.Node) as Group;

                        if (!string.IsNullOrEmpty(MMDHelper.errMsg))
                        {
                            cout("ERROR", MMDHelper.errMsg);
                            break;
                        }

                        txtTermGroupId.Text = tGroup.Id.ToString();
                        txtTermGroupName.Text = tGroup.Name ?? "";
                        txtTermGroupDescr.Text = tGroup.Description ?? "";

                        txtCurSelNode.Text = tGroup.Name;

                        // select tab
                        tabControl1.SelectTab(tabGroup);

                    }
                    else if (e.Node.Level == 2)
                    {
                        // termset
                        TermSet tSet = MMDHelper.GetObj(txtSiteUrl.Text, e.Node.Level, e.Node) as TermSet;

                        if (!string.IsNullOrEmpty(MMDHelper.errMsg))
                        {
                            cout("ERROR", MMDHelper.errMsg);
                            break;
                        }

                        txtTermSetCustomSortOrder.Text = tSet.CustomSortOrder;
                        txtTermSetDescription.Text = tSet.Description;
                        txtTermSetId.Text = tSet.Id.ToString();
                        chkTermSetIsAvailableForTagging.Checked = tSet.IsAvailableForTagging;
                        txtTermSetIsOpenForTermCreation.Text = tSet.IsOpenForTermCreation.ToString();
                        txtTermSetLastModifiedDate.Text = tSet.LastModifiedDate.ToString();
                        txtTermSetName.Text = GenUtil.MmdDenormalize(tSet.Name);

                        txtCurSelNode.Text = tSet.Name;

                        // select tab
                        tabControl1.SelectTab(tabTermSet);

                    }
                    else if (e.Node.Level >= 3)
                    {
                        // term
                        Term term = MMDHelper.GetObj(txtSiteUrl.Text, e.Node.Level, e.Node) as Term;

                        if (!string.IsNullOrEmpty(MMDHelper.errMsg))
                        {
                            cout("ERROR", MMDHelper.errMsg);
                            break;
                        }

                        txtTermCreatedDate.Text = term.CreatedDate.ToString();
                        txtTermCustomSortOrder.Text = term.CustomSortOrder;
                        txtTermId.Text = term.Id.ToString();
                        chkTermIsAvailableForTagging.Checked = term.IsAvailableForTagging;
                        chkTermIsKeyword.Checked = term.IsKeyword;
                        chkTermIsReused.Checked = term.IsReused;
                        chkTermIsRoot.Checked = term.IsRoot;
                        chkTermIsSourceTerm.Checked = term.IsSourceTerm;
                        txtTermLabels.Text = GenUtil.MmdDenormalize(GenUtil.LabelsToString(term.Labels, term.Name));
                        txtTermLastModifiedDate.Text = term.LastModifiedDate.ToString();
                        txtTermName.Text = GenUtil.MmdDenormalize(term.Name);
                        txtTermTermsCount.Text = term.TermsCount.ToString();

                        txtCurSelNode.Text = term.Name;

                        // select tab
                        tabControl1.SelectTab(tabTerm);

                        txtNewTermGuid.Text = NewTermGuidLabel;
                    }

                    break;
                }

            });

            StopWait();
        }


        #region "Utility"

        public void cout()
        {
            txtStatus.Text = Environment.NewLine + txtStatus.Text;
            txtStatus.Refresh();
        }

        public void cout(object o1, object o2)
        {
            if (o1 == null) o1 = "";
            if (o2 == null) o2 = "";
            txtStatus.Text = string.Format("[{0}] {1} : {2}{3}", DateTime.Now.ToLongTimeString(), o1.ToString(), o2.ToString(), Environment.NewLine) + txtStatus.Text;
            txtStatus.Refresh();
        }

        public void cout(object o)
        {
            if (o == null) o = "";
            txtStatus.Text = string.Format("[{0}] {1}{2}", DateTime.Now.ToLongTimeString(), o.ToString(), Environment.NewLine) + txtStatus.Text;
            txtStatus.Refresh();
        }

        private void StartWait()
        {
            picLogoWait.Visible = true;
            picLogoWait.Refresh();
            Cursor = Cursors.WaitCursor;
        }

        private void StopWait()
        {
            picLogoWait.Visible = false;
            picLogoWait.Refresh();
            this.Cursor = Cursors.Default;
        }

        #endregion


        /// <summary>
        /// </summary>
        private void btnExport_Click(object sender, EventArgs e)
        {
            StartWait();

            this.Invoke((MethodInvoker)delegate
            {
                string msg = "";

                if (rbExportUseMS.Checked)
                {
                    if (!MicrosoftExporter.ExportToMsFormat(saveFileDialog1, txtSiteUrl.Text, tvMMD.SelectedNode, out msg))
                    {
                        cout("ERROR", msg);
                    }
                    else
                    {
                        cout("Export Complete.");
                    }

                }
                else if (rbExportUseProp.Checked)
                {
                    if (!ProprietaryExporter.ExportToPropFormat(saveFileDialog1, txtSiteUrl.Text, chkSplitSyns.Checked, tvMMD.SelectedNode, out msg))
                    {
                        cout("ERROR", msg);
                    }
                    else
                    {
                        cout("Export Complete.");
                    }

                }
                else if (rbExportUseXml.Checked)
                {
                    if (!XmlExportHelper.ExportToXml(saveFileDialog2, txtSiteUrl.Text, tvMMD.SelectedNode, out msg))
                    {
                        cout("ERROR", msg);
                    }
                    else
                    {
                        cout("Export Complete.");
                    }

                }

            });

            StopWait();
        }


        /// <summary>
        /// </summary>
        private void btnImport_Click(object sender, EventArgs e)
        {
            StartWait();

            this.Invoke((MethodInvoker)delegate
            {
                while (true)
                {
                    openFileDialog1.FileName = "";
                    openFileDialog1.Filter = "Xml Files|*.xml";

                    if (openFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                        // get current node, which should be group
                        var curNode = tvMMD.SelectedNode;

                        if (curNode == null || curNode.Level != 1)
                        {
                            cout("ERROR", "Please choose a Group in the tree where the termset(s) and term(s) will be imported.");
                            break;
                        }

                        // get group from mmd
                        var tGroup = MMDHelper.GetObj(txtSiteUrl.Text, curNode.Level, curNode) as Group;

                        if (!string.IsNullOrEmpty(MMDHelper.errMsg))
                        {
                            cout("ERROR", MMDHelper.errMsg);
                            break;
                        }

                        // get xml content from file
                        var sr = new StreamReader(openFileDialog1.FileName);
                        string xmlContent = sr.ReadToEnd();
                        sr.Close();

                        var msg = XmlImportHelper.ProcessXml(xmlContent, tGroup, curNode);

                        if (!GenUtil.IsNull(msg))
                        {
                            cout("ERROR", msg);
                            break;
                        }

                        cout("Xml File Imported.");

                        MMDHelper.LoadChildObjects(txtSiteUrl.Text, curNode.Level, curNode);

                        if (!string.IsNullOrEmpty(MMDHelper.errMsg))
                        {
                            cout("ERROR", MMDHelper.errMsg);
                        }

                    }

                    break;
                }

            });

            StopWait();
        }


        #region "Updating"


        /// <summary>
        /// </summary>
        private void btnUpdateTerm_Click(object sender, EventArgs e)
        {
            StartWait();

            this.Invoke((MethodInvoker)delegate
            {

                while(true)
                {
                    TreeNode curNode = tvMMD.SelectedNode;

                    if (curNode == null)
                    {
                        break;
                    }

                    if (txtTermName.Text.Trim().Length == 0)
                    {
                        cout("Cannot update term, name is required.");
                        break;
                    }

                    if (curNode.Level >= 3)
                    {
                        Term term = MMDHelper.GetObj(txtSiteUrl.Text, curNode.Level, curNode) as Term;

                        if (!string.IsNullOrEmpty(MMDHelper.errMsg))
                        {
                            cout("ERROR", MMDHelper.errMsg);
                            break;
                        }

                        try
                        {
                            // term labels
                            string ignoreTermName = GenUtil.MmdNormalize(txtTermName.Text);
                            List<string> lstNewLabels = GenUtil.MmdNormalize(txtTermLabels.Text).Split(new char[] { ';' }).ToList();

                            // delete existing labels (if label in collection not found in textbox)
                            int i = 0;
                            while (i < term.Labels.Select(x => x.Value).ToList().Count())
                            {
                                var curLbl = term.Labels[i];
                                var curLblVal = GenUtil.MmdNormalize(curLbl.Value);

                                if (curLblVal.ToLower() != ignoreTermName.ToLower()
                                    && !lstNewLabels.Any(x => x.Trim().ToLower() == curLblVal.ToLower())
                                    && !curLbl.IsDefaultForLanguage)
                                {
                                    curLbl.Delete();
                                }
                                else
                                {
                                    i++;
                                }
                            }

                            // add new labels (if label in textbox not found in collection)
                            foreach (var lblNew in lstNewLabels)
                            {
                                if (!GenUtil.IsNull(lblNew)
                                    && lblNew.Trim().ToLower() != ignoreTermName.ToLower()
                                    && !term.Labels.Any(x => GenUtil.MmdNormalize(x.Value).ToLower() == lblNew.ToLower()))
                                {
                                    term.CreateLabel(lblNew.Trim(), CultureInfo.CurrentCulture.LCID, false);
                                }
                            }


                            term.IsAvailableForTagging = chkTermIsAvailableForTagging.Checked;
                            term.Name = GenUtil.MmdNormalize(txtTermName.Text);
                            term.TermStore.CommitAll();

                            curNode.Text = string.Format("{1} [{0}]", term.TermsCount, GenUtil.MmdNormalize(txtTermName.Text));

                            cout("Term Updated");

                        }
                        catch (Exception exc)
                        {
                            cout("ERROR updating term", exc.Message);
                        }

                    }

                    break;
                }

            });

            StopWait();
        }


        /// <summary>
        /// </summary>
        private void btnUpdateTermSet_Click(object sender, EventArgs e)
        {
            StartWait();

            this.Invoke((MethodInvoker)delegate
            {
                while(true)
                {
                    TreeNode curNode = tvMMD.SelectedNode;

                    if (curNode == null)
                    {
                        break;
                    }

                    if (txtTermSetName.Text.Trim().Length == 0)
                    {
                        cout("Cannot update termset, name is required.");
                        break;
                    }

                    if (curNode.Level == 2)
                    {
                        TermSet tSet = MMDHelper.GetObj(txtSiteUrl.Text, curNode.Level, curNode) as TermSet;

                        if (!string.IsNullOrEmpty(MMDHelper.errMsg))
                        {
                            cout("ERROR", MMDHelper.errMsg);
                            break;
                        }

                        try
                        {
                            tSet.Name = GenUtil.MmdNormalize(txtTermSetName.Text.Trim());
                            tSet.IsAvailableForTagging = chkTermSetIsAvailableForTagging.Checked;
                            tSet.Description = txtTermSetDescription.Text.Trim();
                            tSet.TermStore.CommitAll();

                            cout("TermSet Updated");

                            curNode.Text = string.Format("{1} [{0}]", tSet.Terms.Count, GenUtil.MmdNormalize(txtTermSetName.Text.Trim()));

                        }
                        catch (Exception exc)
                        {
                            cout("ERROR updating termset", exc.Message);
                        }

                    }
                    
                    break;
                }

            });

            StopWait();
        }


        /// <summary>
        /// </summary>
        private void btnUpdateGroup_Click(object sender, EventArgs e)
        {
            StartWait();

            this.Invoke((MethodInvoker)delegate
            {
                while(true)
                {
                    TreeNode curNode = tvMMD.SelectedNode;

                    if (curNode == null)
                    {
                        break;
                    }

                    if (txtTermGroupName.Text.Trim().Length == 0)
                    {
                        cout("Cannot update group, name is required.");
                        break;
                    }

                    if (curNode.Level == 1)
                    {
                        Group tGroup = MMDHelper.GetObj(txtSiteUrl.Text, curNode.Level, curNode) as Group;

                        if (!string.IsNullOrEmpty(MMDHelper.errMsg))
                        {
                            cout("ERROR", MMDHelper.errMsg);
                            break;
                        }

                        try
                        {
                            tGroup.Description = txtTermGroupDescr.Text.Trim();
                            tGroup.Name = txtTermGroupName.Text.Trim();
                            tGroup.TermStore.CommitAll();

                            curNode.Text = string.Format("{1} [{0}]", tGroup.TermSets.Count, txtTermGroupName.Text.Trim());

                            cout("Term Group Updated");

                        }
                        catch (Exception exc)
                        {
                            cout("ERROR updating term group", exc.Message);
                        }

                    }

                    break;
                }

            });

            StopWait();
        }


        #endregion


        #region "Creating New"


        /// <summary>
        /// </summary>
        private void btnCreateGroup_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Are you sure?", "Confirm", MessageBoxButtons.YesNo) != DialogResult.Yes)
            {
                return;
            }

            StartWait();

            this.Invoke((MethodInvoker)delegate
            {
                while(true)
                {
                    TreeNode curNode = tvMMD.SelectedNode;

                    if (curNode == null)
                    {
                        break;
                    }

                    if (txtTermGroupName.Text.Trim().Length == 0)
                    {
                        cout("Cannot create group, name is required.");
                        break;
                    }

                    TermStore tStore = null;

                    if (curNode.Level == 0)
                    {
                        // create new group in termstore
                        tStore = MMDHelper.GetObj(txtSiteUrl.Text, curNode.Level, curNode) as TermStore;

                        if (!string.IsNullOrEmpty(MMDHelper.errMsg))
                        {
                            cout("ERROR", MMDHelper.errMsg);
                            break;
                        }

                    }
                    else
                    {
                        MessageBox.Show("Cannot create group in current location, select a termstore.");
                        break;
                    }

                    if (tStore != null)
                    {
                        if (tStore.Groups.Any(x => x.Name.Trim().ToLower() == txtTermGroupName.Text.ToLower().Trim()))
                        {
                            cout("Cannot create group, name not unique.");
                            break;
                        }

                        try
                        {
                            // create new group
                            Group newGroup = tStore.CreateGroup(txtTermGroupName.Text.Trim());
                            newGroup.Description = txtTermGroupDescr.Text.Trim();
                            tStore.CommitAll();

                            // add group to tree
                            TreeNode newNode = new TreeNode();
                            newNode.Text = string.Format("{0} [0]", newGroup.Name);
                            newNode.Name = newGroup.Id.ToString();

                            curNode.Nodes.Add(newNode);
                            curNode.Expand();

                            cout("New Group Created");

                        }
                        catch (Exception ex)
                        {
                            cout("ERROR creating group", ex.Message);
                        }
                    }

                    break;
                }

            });

            StopWait();
        }


        /// <summary>
        /// </summary>
        private void btnCreateTermSet_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Are you sure?", "Confirm", MessageBoxButtons.YesNo) != DialogResult.Yes)
            {
                return;
            }

            StartWait();

            this.Invoke((MethodInvoker)delegate
            {
                while(true)
                {
                    TreeNode curNode = tvMMD.SelectedNode;

                    if (curNode == null)
                    {
                        break;
                    }

                    if (txtTermSetName.Text.Trim().Length == 0)
                    {
                        cout("Cannot create termset, name is required.");
                        break;
                    }

                    Group tGroup = null;

                    if (curNode.Level == 1)
                    {
                        // create new termset in group
                        tGroup = MMDHelper.GetObj(txtSiteUrl.Text, curNode.Level, curNode) as Group;

                        if (!string.IsNullOrEmpty(MMDHelper.errMsg))
                        {
                            cout("ERROR", MMDHelper.errMsg);
                            break;
                        }

                    }
                    else
                    {
                        MessageBox.Show("Cannot create termset in current location, select a group.");
                        break;
                    }

                    if (tGroup != null)
                    {
                        if (tGroup.TermSets.Any(x =>
                                GenUtil.MmdNormalize(x.Name).ToLower() == GenUtil.MmdNormalize(txtTermSetName.Text).ToLower()))
                        {
                            cout("Cannot create termset, name not unique.");
                            break;
                        }

                        try
                        {
                            // create new termset
                            TermSet tSet = tGroup.CreateTermSet(GenUtil.MmdNormalize(txtTermSetName.Text.Trim()));
                            tSet.IsAvailableForTagging = chkTermSetIsAvailableForTagging.Checked;
                            tSet.Description = txtTermSetDescription.Text.Trim();
                            tSet.TermStore.CommitAll();

                            // add termset to tree
                            TreeNode newNode = new TreeNode();
                            newNode.Text = string.Format("{0} [0]", GenUtil.MmdDenormalize(tSet.Name));
                            newNode.Name = tSet.Id.ToString();

                            curNode.Nodes.Add(newNode);
                            curNode.Expand();

                            cout("New TermSet Created");

                        }
                        catch (Exception exc)
                        {
                            cout("ERROR creating termset", exc.Message);
                        }

                    }
                    
                    break;
                }
            });

            StopWait();
        }


        /// <summary>
        /// </summary>
        private void btnCreateTerm_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Are you sure?", "Confirm", MessageBoxButtons.YesNo) != DialogResult.Yes)
            {
                return;
            }

            StartWait();

            this.Invoke((MethodInvoker)delegate
            {

                while(true)
                {
                    TreeNode curNode = tvMMD.SelectedNode;

                    if (curNode == null)
                    {
                        break;
                    }

                    if (txtTermName.Text.Trim().Length == 0)
                    {
                        cout("Cannot create term, name is required.");
                        break;
                    }

                    Term newTerm = null;

                    try
                    {
                        if (curNode.Level == 2)
                        {
                            // add term to termset
                            TermSet tSet = MMDHelper.GetObj(txtSiteUrl.Text, curNode.Level, curNode) as TermSet;

                            if (!string.IsNullOrEmpty(MMDHelper.errMsg))
                            {
                                cout("ERROR", MMDHelper.errMsg);
                                break;
                            }

                            if (GenUtil.IsNull(txtNewTermGuid.Text)
                                || txtNewTermGuid.Text.Trim().ToLower() == NewTermGuidLabel.Trim().ToLower()
                                || !GenUtil.IsGuid(txtNewTermGuid.Text))
                            {
                                newTerm = tSet.CreateTerm(GenUtil.MmdNormalize(txtTermName.Text), CultureInfo.CurrentCulture.LCID);
                            }
                            else
                            {
                                if (!GenUtil.IsGuid(txtNewTermGuid.Text))
                                {
                                    MessageBox.Show("Cannot create term, invalid new term Guid.");
                                    break;
                                }
                                else
                                {
                                    newTerm = tSet.CreateTerm(GenUtil.MmdNormalize(txtTermName.Text), CultureInfo.CurrentCulture.LCID, new Guid(txtNewTermGuid.Text));
                                }
                            }

                        }
                        else if (curNode.Level >= 3)
                        {
                            // add term to term
                            Term term = MMDHelper.GetObj(txtSiteUrl.Text, curNode.Level, curNode) as Term;

                            if (!string.IsNullOrEmpty(MMDHelper.errMsg))
                            {
                                cout("ERROR", MMDHelper.errMsg);
                                break;
                            }

                            if (GenUtil.IsNull(txtNewTermGuid.Text)
                                || txtNewTermGuid.Text.Trim().ToLower() == NewTermGuidLabel.Trim().ToLower()
                                || !GenUtil.IsGuid(txtNewTermGuid.Text))
                            {
                                newTerm = term.CreateTerm(GenUtil.MmdNormalize(txtTermName.Text), CultureInfo.CurrentCulture.LCID);
                            }
                            else
                            {
                                if (!GenUtil.IsGuid(txtNewTermGuid.Text))
                                {
                                    MessageBox.Show("Cannot create term, invalid new term Guid.");
                                    break;
                                }
                                else
                                {
                                    newTerm = term.CreateTerm(GenUtil.MmdNormalize(txtTermName.Text), CultureInfo.CurrentCulture.LCID, new Guid(txtNewTermGuid.Text));
                                }
                            }

                        }
                        else
                        {
                            MessageBox.Show("Cannot create term in current location, select a termset or term.");
                            break;
                        }

                        if (newTerm != null)
                        {
                            newTerm.IsAvailableForTagging = chkTermIsAvailableForTagging.Checked;

                            // labels
                            if (!GenUtil.IsNull(txtTermLabels.Text))
                            {
                                foreach (string lbl in GenUtil.MmdNormalize(txtTermLabels.Text).Split(new char[] { ';' }))
                                {
                                    if (!GenUtil.IsNull(lbl)
                                        && lbl.ToLower() != GenUtil.MmdNormalize(newTerm.Name).ToLower())
                                    {
                                        newTerm.CreateLabel(lbl, CultureInfo.CurrentCulture.LCID, false);
                                    }
                                }
                            }

                            newTerm.TermStore.CommitAll();

                            // add term to tree
                            TreeNode newNode = new TreeNode();
                            newNode.Text = string.Format("{0} [0]", GenUtil.MmdDenormalize(newTerm.Name));
                            newNode.Name = newTerm.Id.ToString();

                            curNode.Nodes.Add(newNode);
                            curNode.Expand();

                            cout("New Term Created");

                            txtNewTermGuid.Text = NewTermGuidLabel;

                        }

                    }
                    catch (Exception ex)
                    {
                        cout("ERROR creating term", ex.Message);
                    }

                    break;
                }

            });

            StopWait();
        }


        #endregion


        /// <summary>
        /// </summary>
        private void btnTermDelete_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Are you sure?", "Confirm", MessageBoxButtons.YesNo) != DialogResult.Yes)
            {
                return;
            }

            StartWait();

            this.Invoke((MethodInvoker)delegate
            {

                while (true)
                {
                    TreeNode curNode = tvMMD.SelectedNode;

                    if (curNode == null)
                    {
                        break;
                    }

                    if (curNode.Level < 3)
                    {
                        cout("Can only delete terms.");
                        break;
                    }

                    try
                    {
                        Term term = MMDHelper.GetObj(txtSiteUrl.Text, curNode.Level, curNode) as Term;

                        if (!string.IsNullOrEmpty(MMDHelper.errMsg))
                        {
                            cout("ERROR", MMDHelper.errMsg);
                            break;
                        }

                        if (term == null)
                        {
                            cout("ERROR", "Term not found.");
                            break;
                        }

                        term.Delete();
                        term.TermStore.CommitAll();

                        // delete node from tree
                        curNode.Remove();

                        cout("Term Deleted");

                    }
                    catch (Exception exc)
                    {
                        cout("ERROR deleting term", exc.Message);
                    }

                    break;
                }

            });

            StopWait();
        }


        /// <summary>
        /// </summary>
        private void btnTermSetDelete_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Are you sure?", "Confirm", MessageBoxButtons.YesNo) != DialogResult.Yes)
            {
                return;
            }

            StartWait();

            this.Invoke((MethodInvoker)delegate
            {

                while (true)
                {
                    TreeNode curNode = tvMMD.SelectedNode;

                    if (curNode == null)
                    {
                        break;
                    }

                    if (curNode.Level != 2)
                    {
                        cout("Can only delete termsets.");
                        break;
                    }

                    try
                    {
                        var termSet = MMDHelper.GetObj(txtSiteUrl.Text, curNode.Level, curNode) as TermSet;

                        if (!string.IsNullOrEmpty(MMDHelper.errMsg))
                        {
                            cout("ERROR", MMDHelper.errMsg);
                            break;
                        }

                        if (termSet == null)
                        {
                            cout("ERROR", "Termset not found.");
                            break;
                        }

                        termSet.Delete();
                        termSet.TermStore.CommitAll();

                        // delete node from tree
                        curNode.Remove();

                        cout("Termset Deleted");

                    }
                    catch (Exception exc)
                    {
                        cout("ERROR deleting termset", exc.Message);
                    }

                    break;
                }

            });

            StopWait();
        }


        /// <summary>
        /// </summary>
        private void btnClearTerms_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Are you sure?", "Confirm", MessageBoxButtons.YesNo) != DialogResult.Yes)
            {
                return;
            }

            StartWait();

            this.Invoke((MethodInvoker)delegate
            {

                while (true)
                {
                    TreeNode curNode = tvMMD.SelectedNode;

                    if (curNode == null)
                    {
                        break;
                    }

                    if (curNode.Level >= 3)
                    {
                        cout("Can only clear terms.");
                        break;
                    }

                    try
                    {
                        var term = MMDHelper.GetObj(txtSiteUrl.Text, curNode.Level, curNode) as Term;

                        if (!string.IsNullOrEmpty(MMDHelper.errMsg))
                        {
                            cout("ERROR", MMDHelper.errMsg);
                            break;
                        }

                        if (term == null)
                        {
                            cout("ERROR", "Term not found.");
                            break;
                        }

                        foreach (var curTerm in term.Terms)
                        {
                            DeleteTerms(curTerm);
                        }
                        term.TermStore.CommitAll();

                        curNode.Nodes.Clear();

                        cout("Term Cleared");

                    }
                    catch (Exception exc)
                    {
                        cout("ERROR deleting terms", exc.Message);
                    }

                    break;
                }

            });

            StopWait();
        }


        /// <summary>
        /// </summary>
        private void btnBulkMergeTerms_Click(object sender, EventArgs e)
        {
            StartWait();

            if (tbBulkMergeTerms.Text.IsNull())
            {
                return;
            }

            this.Invoke((MethodInvoker)delegate
            {
                while (true)
                {
                    TreeNode curNode = tvMMD.SelectedNode;

                    if (curNode == null)
                    {
                        break;
                    }

                    if (curNode.Level != 2)
                    {
                        cout("Choose a termset.");
                        break;
                    }

                    try
                    {
                        var termSet = MMDHelper.GetObj(txtSiteUrl.Text, curNode.Level, curNode) as TermSet;

                        if (!string.IsNullOrEmpty(MMDHelper.errMsg))
                        {
                            cout("ERROR", MMDHelper.errMsg);
                            break;
                        }

                        if (termSet == null)
                        {
                            cout("ERROR", "Termset not found.");
                            break;
                        }

                        var newTerms = GenUtil.NormalizeEol(tbBulkMergeTerms.Text.Trim())
                            .Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries)
                            .Where(x => x.Trim().Length > 0)
                            .Distinct()
                            .ToList<string>();

                        bool termsAdded = false;

                        foreach (var newTerm in newTerms)
                        {
                            var termSearchResults = termSet.GetTerms(newTerm.Trim(), false, StringMatchOption.ExactMatch, 1, false);

                            if (termSearchResults.Any())
                            {
                                cout(string.Format("Term found, skipping: {0}", newTerm.Trim()));
                            }
                            else
                            {
                                cout(string.Format("Term not found, creating new term: {0}", newTerm.Trim()));

                                try
                                {
                                    var termCreated = termSet.CreateTerm(newTerm.Trim(), CultureInfo.CurrentCulture.LCID);
                                    termsAdded = true;
                                }
                                catch (Exception ex)
                                {
                                    cout("ERROR creating new term: " + ex.Message);
                                }
                            }
                        }

                        if (termsAdded)
                        {
                            termSet.TermStore.CommitAll();
                        }

                    }
                    catch (Exception exc)
                    {
                        cout("ERROR Bulk Merging Terms", exc.Message);
                    }

                    break; // #important
                }

            });

            StopWait();
        }


        /// <summary>
        /// </summary>
        private void btnTermSetClearTerms_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Are you sure?", "Confirm", MessageBoxButtons.YesNo) != DialogResult.Yes)
            {
                return;
            }

            StartWait();

            this.Invoke((MethodInvoker)delegate
            {

                while (true)
                {
                    TreeNode curNode = tvMMD.SelectedNode;

                    if (curNode == null)
                    {
                        break;
                    }

                    if (curNode.Level != 2)
                    {
                        cout("Can only clear termsets.");
                        break;
                    }

                    try
                    {
                        var termSet = MMDHelper.GetObj(txtSiteUrl.Text, curNode.Level, curNode) as TermSet;

                        if (!string.IsNullOrEmpty(MMDHelper.errMsg))
                        {
                            cout("ERROR", MMDHelper.errMsg);
                            break;
                        }

                        if (termSet == null)
                        {
                            cout("ERROR", "Termset not found.");
                            break;
                        }

                        foreach (var term in termSet.Terms)
                        {
                            DeleteTerms(term);
                        }
                        termSet.TermStore.CommitAll();

                        curNode.Nodes.Clear();

                        cout("Termset Cleared");

                    }
                    catch (Exception exc)
                    {
                        cout("ERROR deleting terms", exc.Message);
                    }

                    break;
                }

            });

            StopWait();
        }


        /// <summary>
        /// </summary>
        private void DeleteTerms(Term term)
        {
            if (term.TermsCount > 0)
            {
                foreach (var subTerm in term.Terms)
                {
                    DeleteTerms(subTerm);
                }
            }

            term.Delete();
        }


        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }


        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (aboutBox1 != null)
                aboutBox1.Close();

            aboutBox1 = new AboutBox1();
            aboutBox1.Show();
        }


        private AboutBox1 aboutBox1 = null;


        protected override void OnClosing(CancelEventArgs e)
        {
            if (aboutBox1 != null)
                aboutBox1.Close();

            base.OnClosing(e);
        }


    }

}
