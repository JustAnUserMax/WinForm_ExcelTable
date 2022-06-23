using ExcelDataReader;
using Microsoft.Office.Interop.Excel;

using System.Data;
using System.Text;

namespace WinForm_ExcelTable
{

    public partial class Form1 : Form
    {
        private string fileName = string.Empty;
        private DataTableCollection _tableCollection = null;
        

        public Form1()
        {
            InitializeComponent();
        }
        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                fileName = openFileDialog1.FileName;
                Text = fileName;
                
                if(fileName != string.Empty)
                {
                    FileStream stream = File.Open(fileName, FileMode.Open, FileAccess.Read);
                    System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                    IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream);
                    DataSet dataSet = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (x) => new ExcelDataTableConfiguration()
                        {
                            UseHeaderRow = true
                        }
                    });
                    
                    _tableCollection = dataSet.Tables;
                    
                    
                    toolStripComboBox1.Items.Clear();
                    foreach (System.Data.DataTable table in _tableCollection)
                    {
                        toolStripComboBox1.Items.Add(table.TableName);
                        dataGridView1.DataSource = table;
                    }
                    toolStripComboBox1.SelectedIndex = 0;
                    
                }
                
            }
            else
            {
                throw new System.Exception("Файл не выбран!");
            }
        }

        

        private void createToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveFileDialog1.CheckFileExists = false;
            DialogResult result = saveFileDialog1.ShowDialog();
            if (DialogResult == DialogResult.OK)
            {
                var app = new Microsoft.Office.Interop.Excel.Application();
                var wb = app.Workbooks.Add();
                wb.SaveAs(saveFileDialog1.FileName);

            }
            saveFileDialog1.CheckFileExists = true;
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();

            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
            //Книга.
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
            //Таблица.
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    ExcelApp.Cells[i + 1, j + 1] = dataGridView1.Rows[i].Cells[j].Value;
                }
            }
            ExcelApp.AlertBeforeOverwriting = false;
            ExcelWorkSheet.ClearArrows();
            string[] file_path = fileName.Split('/');
            ExcelWorkBook.SaveAs($@"E:\Моя папка\Creative\Tables\1{0}", fileName);
            ExcelApp.Quit();
            MessageBox.Show("Завершено успешно!", "Создание новой таблицы",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private void saveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult result = saveFileDialog1.ShowDialog();
        }

        private void importToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void exportToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dataGridView1.AllowUserToAddRows = false;
        }
    }
}