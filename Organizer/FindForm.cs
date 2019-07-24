using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Documents;
using System.Windows.Forms;

namespace Organizer
{
    public partial class FindForm : Form
    {
        TextPointer lastFound = null;
        SpellBox rtbInstance = null;

        public SpellBox RtbInstance
        {
            set { rtbInstance = value; }
            get { return rtbInstance; }
        }

        public string InitialText
        {
            set { txtSearchText.Text = value; }
            get { return txtSearchText.Text; }
        }

        public FindForm()
        {
            InitializeComponent();
            this.TopMost = true;
        }

        private void btnDone_Click(object sender, EventArgs e)
        {
            resetSelection();
            this.Close();
        }

        private void btnFindNext_Click(object sender, EventArgs e)
        {
            resetSelection();

            TextRange index = rtbInstance.Find(txtSearchText.Text, lastFound);
            lastFound = null != index ? index.End : null;
            if (null != index)
            {
                rtbInstance.Parent.Focus();
                rtbInstance.Select(index, Color.LightBlue);
            }
            else
            {
                lastFound = null;
            }

        }

        private void resetSelection()
        {
            rtbInstance.SelectionBackColor = Color.White;
        }

        private void FindForm_Load(object sender, EventArgs e)
        {
            
        }
    }
}
