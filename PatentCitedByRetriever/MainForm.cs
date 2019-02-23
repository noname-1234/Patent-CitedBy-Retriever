using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Windows.Forms;

namespace PatentCitedByRetriever
{
    public partial class MainForm : Form
    {
        private string filePath { get; set; }
        private List<DataTable> dts { get; set; }

        public MainForm()
        {
            InitializeComponent();
        }

        private void btnSelectFile_Click(object sender, System.EventArgs e)
        {
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                string fileContent = string.Empty;

                dialog.InitialDirectory = "C:\\";
                dialog.Filter = "Excel (*.xlsx)|*.xlsx";
                dialog.FilterIndex = 1;
                dialog.RestoreDirectory = true;

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    filePath = dialog.FileName;
                    tbFile.Text = Path.GetFileName(filePath);
                    btnRun.Enabled = true;
                }
            }
        }

        private void loadResultToTabPage()
        {
            foreach (DataTable dt in dts)
            {
                tabControl.TabPages.Add(dt.TableName);
                DataGridView dgv = new DataGridView();
                tabControl.TabPages[tabControl.TabPages.Count - 1].Controls.Add(dgv);

                dgv.Dock = DockStyle.Fill;
                dgv.AllowUserToAddRows = false;
                dgv.DataSource = dt;
            }
        }

        private void btnRun_Click(object sender, System.EventArgs e)
        {
            progressBar.Visible = true;
            backgroundWorker.RunWorkerAsync();
        }

        private void backgroundWorker_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            asyncLog("Start processing");
            asyncLog($"Load data from file {filePath}");

            DataTable dtFile = Utils.getDataTableFromExcel(filePath, 1);
            dts = new List<DataTable>();
            for (int c = 0; c < dtFile.Columns.Count; c++)
            {
                dts.Add(Utils.retrieveDataFromImportedDataTable(dtFile, c, asyncLog));
                int percentage = (c + 1) * 100 / dtFile.Columns.Count;
                backgroundWorker.ReportProgress(percentage);
            }

            asyncLog("Finished!");
        }

        private void asyncLog(string log = "")
        {
            Action action = () => richTextBox.Text += log + "\r\n";
            richTextBox.Invoke(action);
        }

        private void backgroundWorker_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                throw e.Error;
            }

            loadResultToTabPage();

            btnSave.Enabled = true;
            progressBar.Visible = false;
        }

        private void backgroundWorker_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {
            progressBar.Value = e.ProgressPercentage;
        }

        private void richTextBox_TextChanged(object sender, EventArgs e)
        {
            richTextBox.SelectionStart = richTextBox.TextLength;
            richTextBox.ScrollToCaret();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            SaveFileDialog dialog = new SaveFileDialog
            {
                Filter = $"Excel (*.xlsx)|*.xlsx",
                Title = "Save result"
,
                FileName = $"PatentCitedByResult_{DateTime.Now.ToString("yyyyMMddHHmmss")}.xlsx"
            };
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                XLWorkbook wb = new XLWorkbook();
                foreach (DataTable dt in dts)
                {
                    IXLWorksheet ws = wb.AddWorksheet(dt);
                    ws.Table(0).Theme = XLTableTheme.None;
                }
                wb.SaveAs(dialog.FileName);
                MessageBox.Show($"File {dialog.FileName} saved");
            }
        }
    }
}