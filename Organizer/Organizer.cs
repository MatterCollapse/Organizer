using System;
using System.Windows.Forms;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.IO;
using System.Diagnostics;

namespace Organizer
{
    #region ToolStripGroups
    public enum RicherTextBoxToolStripGroups
    {
        FontNameAndSize = 0x2,
        BoldUnderlineItalic = 0x4,
        Alignment = 0x8,
        FontColor = 0x10,
        IndentationAndBullets = 0x20,
        Insert = 0x40,
        Zoom = 0x80
    }
    #endregion //ToolStripGroups

    public partial class Organizer : Form
    {
        string newFileName = "unnamed folder";
        Form newNodeForm = new Form();
        string NOTES_ROOT_PATH = "";
        TreeNode unsavedFile = null;
        bool textChanged = false;
        

        #region Natural Methods
        public Organizer()
        {
            NOTES_ROOT_PATH = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Documents", "Visual Studio Organizer");

            InitializeComponent();
            TreeNode root = PopulateTreeView();
            this.treeView1.NodeMouseClick += new TreeNodeMouseClickEventHandler(this.treeView1_NodeMouseClick);
            this.treeView1.SelectedNode = root;
            loadFile(Path.Combine(getNotesLocation(), this.treeView1.SelectedNode.FullPath, "index.rtf"));
            
        }

        private void SpellBox_Load(object sender, EventArgs e)
        {
            // load system fonts
            foreach (FontFamily family in FontFamily.Families)
            {
                tscmbFont.Items.Add(family.Name);
            }
            tscmbFont.SelectedItem = "Microsoft Sans Serif";

            tscmbFontSize.SelectedItem = "12";

            tstxtZoomFactor.Text = Convert.ToString(spellBox1.ZoomFactor * 100);
            tsbtnWordWrap.Checked = spellBox1.WordWrap;
        }

        String getNotesLocation()
        {
            return NOTES_ROOT_PATH;
        }

        private void timer1_Tick_1(object sender, EventArgs e)
        {
            saveFile(treeView1.SelectedNode.FullPath);
        }

        private void renameConfirmClickHandler(object sender, MouseEventArgs e)
        {
            String nodePath = treeView1.SelectedNode.FullPath;
            createFile(Path.Combine(getNotesLocation(), nodePath, newFileName));
            newNodeForm.Close();
        }
        
        private void changeTextChanged(object sender, EventArgs e)
        {
            TextBox textBox = sender as TextBox;
            if (textBox != null)
            {
                newFileName = textBox.Text;
            }
        }
        
        private void tsbtnCreate_Click(object sender, EventArgs e)
        {
            Button button1 = new Button();
            Button button2 = new Button();

            newNodeForm.Size = new Size(250, 150);
            TextBox renamebox = new TextBox();
            renamebox.Location = new Point(newNodeForm.Left + (newNodeForm.Width / 2) - (renamebox.Width / 2) - 10, 10);
            renamebox.TextChanged += new EventHandler(this.changeTextChanged);
            button1.Text = "Confirm";
            button1.Location = new Point(newNodeForm.Right - button1.Width - 25, renamebox.Height + renamebox.Top + 10);
            button1.MouseClick += new MouseEventHandler(this.renameConfirmClickHandler);

            button2.Text = "Cancel";
            button2.Location = new Point(newNodeForm.Left + 10, button1.Top);
            newNodeForm.Text = "Name Folder";
            newNodeForm.HelpButton = true;

            newNodeForm.FormBorderStyle = FormBorderStyle.FixedDialog;
            newNodeForm.AcceptButton = button1;
            newNodeForm.CancelButton = button2;
            newNodeForm.StartPosition = FormStartPosition.CenterParent; //FormStartPosition.CenterScreen;

            newNodeForm.Controls.Add(renamebox);
            newNodeForm.Controls.Add(button1);
            newNodeForm.Controls.Add(button2);

            newNodeForm.ShowDialog();
        }

        private void tsbtnLoad_Click(object sender, EventArgs e)
        {
            saveFile(treeView1.SelectedNode.FullPath);

            using (var fbd = new FolderBrowserDialog()) //This evokes filebrowser, which (unlike OpenFileDialogue) allows selection of folders
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    Debug.Print("Loading Path: " + fbd.SelectedPath);

                    NOTES_ROOT_PATH = fbd.SelectedPath; //This is our new root

                    //re-innitialize the tree
                    this.treeView1.Nodes.Clear();
                    TreeNode root = PopulateTreeView();
                    this.treeView1.SelectedNode = root;
                    loadFile(Path.Combine(getNotesLocation(), this.treeView1.SelectedNode.FullPath, "index.rtf"));
                }
            }
        }

        private void tsbtnDelete_Click(object sender, EventArgs e)
        {
            deleteFile(this.treeView1.SelectedNode.FullPath);
        }

        private void spellBox1_ChildChanged(object sender, System.Windows.Forms.Integration.ChildChangedEventArgs e)
        {
            textChanged = true;
        }
        #endregion //Natural Methods

        #region Node Tree Methods
        private TreeNode PopulateTreeView()
        {
            TreeNode rootNode;

            DirectoryInfo info = new DirectoryInfo(Path.Combine(getNotesLocation(), "Notes"));
            if (!info.Exists)
                createNoteFolder(info.FullName);

            rootNode = new TreeNode(info.Name);
            unsavedFile = rootNode;
            rootNode.Tag = info;
            GetDirectories(info.GetDirectories(), rootNode);
            //GetDirectories(info.GetFiles("*.*"), rootNode); TODO check for files with links
            treeView1.Nodes.Add(rootNode);
            return rootNode;
        }

        private void GetDirectories(DirectoryInfo[] subDirs, TreeNode nodeToAddTo)
        {
            TreeNode aNode;
            DirectoryInfo[] subSubDirs;
            foreach (DirectoryInfo subDir in subDirs)
            {
                aNode = new TreeNode(subDir.Name, 0, 0);
                aNode.Tag = subDir;
                aNode.ImageKey = "folder";
                subSubDirs = subDir.GetDirectories();

                GetDirectories(subSubDirs, aNode); //Removed if dir != 0, worked

                nodeToAddTo.Nodes.Add(aNode);
            }
        }

        /*private void GetDirectories(FileInfo[] subDirs, TreeNode nodeToAddTo)
        {
            TreeNode aNode;
            FileInfo[] subSubDirs;
            foreach (FileInfo subDir in subDirs)
            {
                aNode = new TreeNode(subDir.Name, 0, 0);
                aNode.Tag = subDir;
                aNode.ImageKey = "file";
                nodeToAddTo.Nodes.Add(aNode);
            }
        }*/

        void treeView1_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            TreeNode newSelected = e.Node;
            treeView1.SelectedNode = e.Node;
            ActiveNode.Text = treeView1.SelectedNode.Text;

            if (unsavedFile != newSelected)
            {
                Debug.Print("New Selection");
                saveFile(unsavedFile.FullPath);
                unsavedFile = newSelected;

            }
            else
            {
                Debug.Print("Old Selection");
                saveFile(unsavedFile.FullPath);
                
            }

            spellBox1.Clear();
            DirectoryInfo nodeDirInfo = newSelected.Tag as DirectoryInfo;
            loadFile(Path.Combine(getNotesLocation(), nodeDirInfo.FullName, "index.rtf"));
        }
        #endregion //Node Tree Methods

        #region My Methods
        void loadFile(string filePath)
        {
            FileInfo file = new FileInfo(filePath);
            if (file.Exists && file.Length > 0)
                spellBox1.LoadFile(filePath); //The index is written in the textbox as an rtf
            ActiveNode.Text = treeView1.SelectedNode.Text;
        }

        void createFile(string filePath)
        {
            if (Directory.Exists(filePath))
            {
                Debug.Print("Attempted to create existing file");
            }
            else
            {
                DirectoryInfo[] dir = new DirectoryInfo[] { createNoteFolder(filePath) };
                GetDirectories(dir, treeView1.SelectedNode);
                saveFile(treeView1.SelectedNode.FullPath);
                loadFile(Path.Combine(filePath, "index.rtf"));
                ActiveNode.Text = newFileName;
            }

        }

        DirectoryInfo createNoteFolder(string filePath)
        {
            string index = Path.Combine(filePath, "index.rtf");

            DirectoryInfo dir = Directory.CreateDirectory(filePath);
            FileStream file = File.Create(index);
            file.Close();
            return dir;
        }

        void saveFile(string filePath)
        {
            if (textChanged)
            {
                //save
                Debug.Print("Saving: " + filePath);
                DirectoryInfo dir = unsavedFile.Tag as DirectoryInfo;
                spellBox1.SaveFile(Path.Combine(getNotesLocation(), filePath, "index.rtf"));
                textChanged = false;

                //Identify links: match words to node names using GetDirectories. Create links. On click, links find node with matching name and open it

            }
        }

        //string searchNodes(string name)
        //{

            //GetDirectories();
            //return 
        //}

        void deleteFile(string filePath)
        {
            FileInfo file = new FileInfo(Path.Combine(getNotesLocation(), filePath, "index.rtf"));
            DirectoryInfo folder = new DirectoryInfo(Path.Combine(getNotesLocation(), filePath));

            if (folder.Exists && file.Exists)
            {
                DialogResult result = MessageBox.Show("Delete Selected?", "!warning!", MessageBoxButtons.OKCancel);
                if (result == DialogResult.OK)
                {
                    Debug.Print("Deleting: " + folder.FullName);
                    folder.Delete(true);

                    this.treeView1.Nodes.Clear();
                    TreeNode root = PopulateTreeView();
                    this.treeView1.SelectedNode = root;
                    loadFile(Path.Combine(getNotesLocation(), this.treeView1.SelectedNode.FullPath, "index.rtf"));
                }
            }
            else
            {
                MessageBox.Show("Select a folder first");
            }
        }
        #endregion //My Methods

        #region Settings
        private int indent = 50;
        [Category("Settings")]
        [Description("Value indicating the number of characters used for indentation")]
        public int INDENT
        {
            get { return indent; }
            set { indent = value; }
        }
        #endregion

        #region Toolstrip items handling

        private void tsbtnBIU_Click(object sender, EventArgs e)
        {
            // bold, italic, underline
            try
            {
                if (!(spellBox1.Font == null))
                {
                    Font currentFont = spellBox1.Font;
                    FontStyle newFontStyle = spellBox1.Font.Style;
                    string txt = (sender as ToolStripButton).Name;
                    if (txt.IndexOf("Bold") >= 0)
                        newFontStyle = spellBox1.Font.Style ^ FontStyle.Bold;
                    else if (txt.IndexOf("Italic") >= 0)
                        newFontStyle = spellBox1.Font.Style ^ FontStyle.Italic;
                    else if (txt.IndexOf("Underline") >= 0)
                        newFontStyle = spellBox1.Font.Style ^ FontStyle.Underline;

                    spellBox1.Font = new Font(currentFont.FontFamily, currentFont.Size, newFontStyle);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }

        private void spellBox1_SelectionChanged(object sender, EventArgs e)
        {
            if (spellBox1.Font != null)
            {
                tsbtnBold.Checked = spellBox1.Font.Bold;
                tsbtnItalic.Checked = spellBox1.Font.Italic;
                tsbtnUnderline.Checked = spellBox1.Font.Underline;

                boldToolStripMenuItem.Checked = spellBox1.Font.Bold;
                italicToolStripMenuItem.Checked = spellBox1.Font.Italic;
                underlineToolStripMenuItem.Checked = spellBox1.Font.Underline;

                switch (spellBox1.SelectionAlignment)
                {
                    case System.Windows.HorizontalAlignment.Left:
                        tsbtnAlignLeft.Checked = true;
                        tsbtnAlignCenter.Checked = false;
                        tsbtnAlignRight.Checked = false;

                        leftToolStripMenuItem.Checked = true;
                        centerToolStripMenuItem.Checked = false;
                        rightToolStripMenuItem.Checked = false;
                        break;

                    case System.Windows.HorizontalAlignment.Center:
                        tsbtnAlignLeft.Checked = false;
                        tsbtnAlignCenter.Checked = true;
                        tsbtnAlignRight.Checked = false;

                        leftToolStripMenuItem.Checked = false;
                        centerToolStripMenuItem.Checked = true;
                        rightToolStripMenuItem.Checked = false;
                        break;

                    case System.Windows.HorizontalAlignment.Right:
                        tsbtnAlignLeft.Checked = false;
                        tsbtnAlignCenter.Checked = false;
                        tsbtnAlignRight.Checked = true;

                        leftToolStripMenuItem.Checked = false;
                        centerToolStripMenuItem.Checked = false;
                        rightToolStripMenuItem.Checked = true;
                        break;
                }

                tsbtnBullets.Checked = spellBox1.SelectionBullet;
                bulletsToolStripMenuItem.Checked = spellBox1.SelectionBullet;

                tscmbFont.SelectedItem = spellBox1.Font.FontFamily.Name;
                tscmbFontSize.SelectedItem = spellBox1.Font.Size.ToString();
            }
        }

        private void tsbtnAlignment_Click(object sender, EventArgs e)
        {
            // alignment: left, center, right
            try
            {
                string txt = (sender as ToolStripButton).Name;
                if (txt.IndexOf("Left") >= 0)
                {
                    spellBox1.SelectionAlignment = System.Windows.HorizontalAlignment.Left;
                    tsbtnAlignLeft.Checked = true;
                    tsbtnAlignCenter.Checked = false;
                    tsbtnAlignRight.Checked = false;
                }
                else if (txt.IndexOf("Center") >= 0)
                {
                    spellBox1.SelectionAlignment = System.Windows.HorizontalAlignment.Center;
                    tsbtnAlignLeft.Checked = false;
                    tsbtnAlignCenter.Checked = true;
                    tsbtnAlignRight.Checked = false;
                }
                else if (txt.IndexOf("Right") >= 0)
                {
                    spellBox1.SelectionAlignment = System.Windows.HorizontalAlignment.Right;
                    tsbtnAlignLeft.Checked = false;
                    tsbtnAlignCenter.Checked = false;
                    tsbtnAlignRight.Checked = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }

        private void tsbtnFontColor_Click(object sender, EventArgs e)
        {
            // font color
            try
            {
                using (ColorDialog dlg = new ColorDialog())
                {
                    dlg.Color = spellBox1.ForeColor;
                    if (dlg.ShowDialog() == DialogResult.OK)
                    {
                        spellBox1.ForeColor = dlg.Color;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }

        private void tsbtnBulletsAndNumbering_Click(object sender, EventArgs e)
        {
            // bullets, indentation
            try
            {
                string name = (sender as ToolStripButton).Name;
                if (name.IndexOf("Bullets") >= 0)
                    spellBox1.SelectionBullet = tsbtnBullets.Checked;
                else if (name.IndexOf("Indent") >= 0)
                    spellBox1.SelectionIndent += INDENT;
                else if (name.IndexOf("Outdent") >= 0)
                    spellBox1.SelectionIndent -= INDENT;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }

        private void tscmbFontSize_Click(object sender, EventArgs e)
        {
            // font size
            try
            {
                if (!(spellBox1.Font == null))
                {
                    Font currentFont = spellBox1.Font;
                    float newSize = Convert.ToSingle(tscmbFontSize.SelectedItem.ToString());
                    spellBox1.Font = new Font(currentFont.FontFamily, newSize, currentFont.Style);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }


        private void tscmbFontSize_TextChanged(object sender, EventArgs e)
        {
            // font size custom
            try
            {
                if (!(spellBox1.Font == null))
                {
                    Font currentFont = spellBox1.Font;
                    float newSize = Convert.ToSingle(tscmbFontSize.Text);
                    spellBox1.Font = new Font(currentFont.FontFamily, newSize, currentFont.Style);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }

        private void tscmbFont_Click(object sender, EventArgs e)
        {
            // font
            try
            {
                if (!(spellBox1.Font == null))
                {
                    Font currentFont = spellBox1.Font;
                    FontFamily newFamily = new FontFamily(tscmbFont.SelectedItem.ToString());
                    spellBox1.Font = new Font(newFamily, currentFont.Size, currentFont.Style);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }

        private void btnChooseFont_Click(object sender, EventArgs e)
        {
            using (FontDialog dlg = new FontDialog())
            {
                if (spellBox1.Font != null) dlg.Font = spellBox1.Font;
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    spellBox1.Font = dlg.Font;
                }
            }
        }

        private void tsbtnInsertPicture_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog dlg = new OpenFileDialog())
            {
                dlg.Title = "Insert image";
                dlg.DefaultExt = "png";
                dlg.Filter = "PNG Files|*.png|JPEG Files|*.jpg|Bitmap Files|*.bmp|GIF Files|*.gif|All files|*.*";
                dlg.FilterIndex = 1;
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        string strImagePath = dlg.FileName;
                        Image img = Image.FromFile(strImagePath);
                        Clipboard.SetDataObject(img);
                        this.spellBox1.Paste();
                    }
                    catch
                    {
                        MessageBox.Show("Unable to insert image.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void tsbtnZoomIn_Click(object sender, EventArgs e)
        {
            if (spellBox1.ZoomFactor < 64.0f - 0.20f)
            {
                spellBox1.ZoomFactor += 0.20f;
                tstxtZoomFactor.Text = String.Format("{0:F0}", spellBox1.ZoomFactor * 100);
            }
        }

        private void tsbtnZoomOut_Click(object sender, EventArgs e)
        {
            if (spellBox1.ZoomFactor > 0.16f + 0.20f)
            {
                spellBox1.ZoomFactor -= 0.20f;
                tstxtZoomFactor.Text = String.Format("{0:F0}", spellBox1.ZoomFactor * 100);
            }
        }

        private void tstxtZoomFactor_Leave(object sender, EventArgs e)
        {
            try
            {
                spellBox1.ZoomFactor = Convert.ToSingle(tstxtZoomFactor.Text) / 100;
            }
            catch (FormatException)
            {
                MessageBox.Show("Enter valid number", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                tstxtZoomFactor.Focus();
                tstxtZoomFactor.SelectAll();
            }
            catch (OverflowException)
            {
                MessageBox.Show("Enter valid number", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                tstxtZoomFactor.Focus();
                tstxtZoomFactor.SelectAll();
            }
            catch (ArgumentException)
            {
                MessageBox.Show("Zoom factor should be between 20% and 6400%", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                tstxtZoomFactor.Focus();
                tstxtZoomFactor.SelectAll();
            }
        }


        private void tsbtnWordWrap_Click(object sender, EventArgs e)
        {
            spellBox1.WordWrap = tsbtnWordWrap.Checked;
        }

        #endregion

        #region Changing visibility of toolstrip items

        public void HideToolstripItemsByGroup(RicherTextBoxToolStripGroups group, bool visible)
        {
            if ((group & RicherTextBoxToolStripGroups.FontNameAndSize) != 0)
            {
                tscmbFont.Visible = visible;
                tscmbFontSize.Visible = visible;
                tsbtnChooseFont.Visible = visible;
                toolStripSeparator1.Visible = visible;
            }
            if ((group & RicherTextBoxToolStripGroups.BoldUnderlineItalic) != 0)
            {
                tsbtnBold.Visible = visible;
                tsbtnItalic.Visible = visible;
                tsbtnUnderline.Visible = visible;
                toolStripSeparator2.Visible = visible;
            }
            if ((group & RicherTextBoxToolStripGroups.Alignment) != 0)
            {
                tsbtnAlignLeft.Visible = visible;
                tsbtnAlignRight.Visible = visible;
                tsbtnAlignCenter.Visible = visible;
                toolStripSeparator3.Visible = visible;
            }
            if ((group & RicherTextBoxToolStripGroups.FontColor) != 0)
            {
                tsbtnFontColor.Visible = visible;
                tsbtnWordWrap.Visible = visible;
                toolStripSeparator4.Visible = visible;
            }
            if ((group & RicherTextBoxToolStripGroups.IndentationAndBullets) != 0)
            {
                tsbtnIndent.Visible = visible;
                tsbtnOutdent.Visible = visible;
                tsbtnBullets.Visible = visible;
                toolStripSeparator5.Visible = visible;
            }
            if ((group & RicherTextBoxToolStripGroups.Insert) != 0)
            {
                tsbtnInsertPicture.Visible = visible;
                toolStripSeparator7.Visible = visible;
            }
            if ((group & RicherTextBoxToolStripGroups.Zoom) != 0)
            {
                tsbtnZoomOut.Visible = visible;
                tsbtnZoomIn.Visible = visible;
                tstxtZoomFactor.Visible = visible;
            }
        }

        public bool IsGroupVisible(RicherTextBoxToolStripGroups group)
        {
            switch (group)
            {

                case RicherTextBoxToolStripGroups.FontNameAndSize:
                    return tscmbFont.Visible && tscmbFontSize.Visible && tsbtnChooseFont.Visible && toolStripSeparator1.Visible;

                case RicherTextBoxToolStripGroups.BoldUnderlineItalic:
                    return tsbtnBold.Visible && tsbtnItalic.Visible && tsbtnUnderline.Visible && toolStripSeparator2.Visible;

                case RicherTextBoxToolStripGroups.Alignment:
                    return tsbtnAlignLeft.Visible && tsbtnAlignRight.Visible && tsbtnAlignCenter.Visible && toolStripSeparator3.Visible;

                case RicherTextBoxToolStripGroups.FontColor:
                    return tsbtnFontColor.Visible && tsbtnWordWrap.Visible && toolStripSeparator4.Visible;

                case RicherTextBoxToolStripGroups.IndentationAndBullets:
                    return tsbtnIndent.Visible && tsbtnOutdent.Visible && tsbtnBullets.Visible && toolStripSeparator5.Visible;

                case RicherTextBoxToolStripGroups.Insert:
                    return tsbtnInsertPicture.Visible && toolStripSeparator7.Visible;

                case RicherTextBoxToolStripGroups.Zoom:
                    return tsbtnZoomOut.Visible && tsbtnZoomIn.Visible && tstxtZoomFactor.Visible;

                default:
                    return false;
            }
        }
        #endregion

        #region Public methods for accessing the functionality of the RicherTextBox

        public void SetFontFamily(FontFamily family)
        {
            if (family != null)
            {
                tscmbFont.SelectedItem = family.Name;
            }
        }

        public void SetFontSize(float newSize)
        {
            tscmbFontSize.Text = newSize.ToString();
        }

        public void ToggleBold()
        {
            tsbtnBold.PerformClick();
        }

        public void ToggleItalic()
        {
            tsbtnItalic.PerformClick();
        }

        public void ToggleUnderline()
        {
            tsbtnUnderline.PerformClick();
        }

        public void SetAlign(System.Windows.HorizontalAlignment alignment)
        {
            switch (alignment)
            {
                case System.Windows.HorizontalAlignment.Center:
                    tsbtnAlignCenter.PerformClick();
                    break;

                case System.Windows.HorizontalAlignment.Left:
                    tsbtnAlignLeft.PerformClick();
                    break;

                case System.Windows.HorizontalAlignment.Right:
                    tsbtnAlignRight.PerformClick();
                    break;
            }
        }

        public void Indent()
        {
            tsbtnIndent.PerformClick();
        }

        public void Outdent()
        {
            tsbtnOutdent.PerformClick();
        }

        public void ToggleBullets()
        {
            tsbtnBullets.PerformClick();
        }

        public void ZoomIn()
        {
            tsbtnZoomIn.PerformClick();
        }

        public void ZoomOut()
        {
            tsbtnZoomOut.PerformClick();
        }

        public void ZoomTo(float factor)
        {
            spellBox1.ZoomFactor = factor;
        }

        public void SetWordWrap(bool activated)
        {
            spellBox1.WordWrap = activated;
        }
        #endregion

        #region Context menu handlers
        private void cutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            spellBox1.Cut();
        }

        private void copyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            spellBox1.Copy();
        }

        private void pasteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            spellBox1.Paste();
        }

        private void deleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            spellBox1.Clear();
        }

        private void selectAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            spellBox1.SelectAll();
        }

        private void undoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            spellBox1.Undo();
        }

        private void redoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            spellBox1.Redo();
        }

        private void leftToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tsbtnAlignLeft.PerformClick();

            leftToolStripMenuItem.Checked = true;
            centerToolStripMenuItem.Checked = false;
            rightToolStripMenuItem.Checked = false;


        }

        private void centerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tsbtnAlignCenter.PerformClick();

            leftToolStripMenuItem.Checked = false;
            centerToolStripMenuItem.Checked = true;
            rightToolStripMenuItem.Checked = false;
        }

        private void rightToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tsbtnAlignRight.PerformClick();

            leftToolStripMenuItem.Checked = false;
            centerToolStripMenuItem.Checked = false;
            rightToolStripMenuItem.Checked = true;
        }

        private void boldToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tsbtnBold.PerformClick();
        }

        private void italicToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tsbtnItalic.PerformClick();
        }

        private void underlineToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tsbtnUnderline.PerformClick();
        }

        private void increaseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tsbtnIndent.PerformClick();
        }

        private void decreaseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tsbtnOutdent.PerformClick();
        }

        private void bulletsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tsbtnBullets.PerformClick();
        }

        private void zoomInToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tsbtnZoomIn.PerformClick();
        }

        private void zoomOuToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tsbtnZoomOut.PerformClick();
        }

        private void insertPictureToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tsbtnInsertPicture.PerformClick();
        }

        #endregion

        #region Find and Replace
        private void tsbtnFind_Click(object sender, EventArgs e)
        {
            FindForm findForm = new FindForm();
            findForm.RtbInstance = this.spellBox1;
            findForm.InitialText = "";
            findForm.Show();
        }

        private void tsbtnReplace_Click(object sender, EventArgs e)
        {
            ReplaceForm replaceForm = new ReplaceForm();
            replaceForm.RtbInstance = this.spellBox1;
            replaceForm.InitialText = "";
            replaceForm.Show();
        }

        #endregion
        
    }
}
