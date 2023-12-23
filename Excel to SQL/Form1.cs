
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using System.IO;

namespace Excel_to_SQL
{
    public partial class Form1 : Form
    {
        public string filePath = "";
        public string strOutput = "";
        public Form1()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        }

        private void btnUpload_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            dialog.Title = "Select Excel File";
            dialog.Filter = "Excel Files|*.xlsx|All Files|*.*";

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                filePath = dialog.FileName;
                ExcelReader reader = new ExcelReader();
                List<string> sheets = reader.GetSheetNames(filePath);

                btnUpload.Text = Path.GetFileName(filePath);
                foreach (string sheet in sheets)
                {
                    cbSheet.Items.Add(sheet);
                }
            }
        }

        private void cbSheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            ExcelReader eReader = new ExcelReader();
            List<string> cols = eReader.GetColumnNames(filePath, cbSheet.Text);

            foreach (string col in cols)
            {
                cbDate.Items.Add(col);
                cbConsumer.Items.Add(col);
                cbRcpt.Items.Add(col);
                cbCash.Items.Add(col);
                cbCheque.Items.Add(col);
            }
        }
        
        private void btnStart_Click(object sender, EventArgs e)
        {
            //error handling
            if (filePath == "")
            {
                MessageBox.Show("File path is null", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            int[] cols = { cbDate.SelectedIndex, cbConsumer.SelectedIndex, cbRcpt.SelectedIndex, cbCash.SelectedIndex, cbCheque.SelectedIndex };
            ExcelReader r = new ExcelReader();
            txtOutput.Text =  r.GenerateInsertQueries(filePath, cols, cbSheet.Text, txtKiosk.Text, txtOffice.Text);
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            filePath = string.Empty;
            strOutput = string.Empty;
            btnUpload.Text = "Upload File";

            ClearUI(this);
        }

        private void ClearUI(Control parentControl)
        {
            // Iterate through all controls within the specified parent control
            foreach (Control control in parentControl.Controls)
            {
                // Recursively clear ComboBox items if the control is a container control
                if (control.HasChildren)
                {
                    ClearUI(control);
                }

                // Check if the control is a ComboBox
                if (control is ComboBox comboBox)
                {
                    // Clear the items in the ComboBox
                    comboBox.Items.Clear();
                }
                else if(control is TextBox tb)
                {
                    tb.Text = string.Empty;
                }else if(control is RichTextBox richTextBox)
                {
                    richTextBox.Text = string.Empty;
                }
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            dialog.FileName = txtOffice.Text+".txt";
            dialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";

            if(dialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = dialog.FileName;
                try
                {
                    File.WriteAllText(filePath, txtOutput.Text);
                    MessageBox.Show("File saved successfully!", "Save As", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch(Exception ex)
                {
                    MessageBox.Show($"Error : {ex.Message}", "Save As", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void btnCopy_Click(object sender, EventArgs e)
        {
            try
            {
                Clipboard.SetText(txtOutput.Text);
            }
            catch
            {

            }
        }
    }
}
