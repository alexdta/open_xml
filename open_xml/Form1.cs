using System;
using System.Windows.Forms;

namespace open_xml
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                office_word.editarDoc();

                Close();
            }
            catch (Exception lExcp)
            {
                MessageBox.Show(lExcp.Message);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                office_excel.editarXLSX();

                Close();
            }
            catch (Exception lExcp)
            {
                MessageBox.Show(lExcp.Message);
            }
        }
    }
}