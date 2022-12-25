using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDataReader;
using Word = Microsoft.Office.Interop.Word;

namespace readExcel
{
    public partial class Form1 : Form
    {
        private string fileName = string.Empty;
        private DataTableCollection tableCollection = null;
        public Form1()
        {
            InitializeComponent();
        }

        // открываем файл Excel
        private void TSMI_Open_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult res = openFileDialog1.ShowDialog();
                if (res == DialogResult.OK)
                {
                    fileName = openFileDialog1.FileName;
                    Text = fileName;

                    OpenExcelFile(fileName);
                }
                else
                {
                    throw new Exception("Файл не выбран!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
     
        private void OpenExcelFile(string patch)
        {
            //System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            FileStream stream = File.Open(patch, FileMode.Open, FileAccess.Read);

            IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream);

            DataSet db = reader.AsDataSet(new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = (x) => new ExcelDataTableConfiguration()
                {
                    UseHeaderRow = true
                }
            });
            tableCollection = db.Tables;
            toolStripComboBox1.Items.Clear();
            foreach (DataTable tabe in tableCollection)
            {
                toolStripComboBox1.Items.Add(tabe.TableName);
            }
            toolStripComboBox1.SelectedIndex = 0;

        }
        private void toolStripComboBox1_SelectedIndexChanged(object sender, EventArgs e) // загрузка данных из excel в dataGridView
        {
            DataTable table = tableCollection[Convert.ToString(toolStripComboBox1.SelectedItem)];
            dataGridView1.DataSource = table;
        }
       

        private void tsmiHelp_Click(object sender, EventArgs e)//вызов справки 
        {
            Help.ShowHelp(this, "Help.chm");
        }

        private void Form1_Load(object sender, EventArgs e) 
        {
            dataGridView2.RowCount = 3; // загрузка количества колонок

        }

        private void btnSumm_Click(object sender, EventArgs e) // суммирование колонок dataGridView
        {
            //dataGridView2.CurrentRow.Cells[1].Value.ToString() == dataGridView1.CurrentRow.Cells[1].Value.ToString();
            /*dataGridView2.Rows[0].Cells[1].Value = "hhgh"*/
            /*dataGridView1.Rows[3].Cells[3];*/

            // dataGridView2.Rows[0].Cells[0].Value = dataGridView1.CurrentRow.Cells[3].Value.ToString(); перенос

            int lec = 0;
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {

                lec = lec + Convert.ToInt32(dataGridView1["лек", i].Value);
                dataGridView2.Rows[0].Cells[0].Value = lec.ToString();
            }

            int pra = 0;
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {

                pra = pra + Convert.ToInt32(dataGridView1["пра", i].Value);
                dataGridView2.Rows[0].Cells[1].Value = pra.ToString();
            }

            int lab = 0;
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {

                lab = lab + Convert.ToInt32(dataGridView1["лаб", i].Value);
                dataGridView2.Rows[0].Cells[2].Value = lab.ToString();
            }

            int zach = 0;
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {

                zach = zach + Convert.ToInt32(dataGridView1["зач", i].Value);
                dataGridView2.Rows[0].Cells[3].Value = zach.ToString();
            }

            int aykz = 0;
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {

                aykz += Convert.ToInt32(dataGridView1["экз", i].Value);
                dataGridView2.Rows[0].Cells[4].Value = aykz.ToString();
            }

            int kons = 0;
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {

                kons += Convert.ToInt32(dataGridView1["КОНС", i].Value);
                dataGridView2.Rows[0].Cells[5].Value = kons.ToString();
            }

            int praY = 0;
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {

                praY += Convert.ToInt32(dataGridView1["ПРА_1", i].Value);
                dataGridView2.Rows[0].Cells[6].Value = praY.ToString();
            }
            int summa = 0;
            summa = praY + kons + aykz + lab + pra + lec + zach;

            dataGridView2.Rows[0].Cells[11].Value = summa.ToString();

            //foreach(DataGridViewRow r in dataGridView1.Rows)
            //{
            //    sum += (Double)r.Cells[0].Value;
            //    }
            //decimal i = 0;
            //for (Int32 j = 0; j < dataGridView1.Rows.Count; j++)
            //{
            //    i += Convert.ToDecimal(dataGridView1.Rows[j].Cells[10].Value);
            //}

        }


        //private void ToCsV(DataGridView dGV, string filename)
        //{
        //    string stOutput = "";
        //    // Export titles:
        //    string sHeaders = "";

        //    for (int j = 0; j < dGV.Columns.Count; j++)
        //        sHeaders = sHeaders.ToString() + Convert.ToString(dGV.Columns[j].HeaderText) + "\t";
        //    stOutput += sHeaders + "\r\n";
        //    // Export data.
        //    for (int i = 0; i < dGV.RowCount - 1; i++)
        //    {
        //        string stLine = "";
        //        for (int j = 0; j < dGV.Rows[i].Cells.Count; j++)
        //            stLine = stLine.ToString() + Convert.ToString(dGV.Rows[i].Cells[j].Value) + "\t";
        //        stOutput += stLine + "\r\n";
        //    }
        //    Encoding utf16 = Encoding.GetEncoding(1254);
        //    byte[] output = utf16.GetBytes(stOutput);
        //    FileStream fs = new FileStream(filename, FileMode.Create);
        //    BinaryWriter bw = new BinaryWriter(fs);
        //    bw.Write(output, 0, output.Length); //write the encoded file
        //    bw.Flush();
        //    bw.Close();
        //    fs.Close();


        //    int iRowCount = 11;

        //    for (int i = 0; i < dataGridView2.Rows.Count; i++)
        //    {
        //        wSheet.Cells[iRowCount, 1] = dataGridView2.Rows[i].Cells[0].Value.ToString();
        //        wSheet.Cells[iRowCount, 2] = dataGridView2.Rows[i].Cells[1].Value.ToString();
        //        wSheet.Cells[iRowCount, 3] = dataGridView2.Rows[i].Cells[2].Value.ToString();
        //        wSheet.Cells[iRowCount, 4] = dataGridView2.Rows[i].Cells[3].Value.ToString();

        //        iRowCount++;

        //        // Добавляем строчку ниже
        //        var cellsDRnr = wSheet.get_Range("A" + iRowCount, "A" + iRowCount);
        //        cellsDRnr.EntireRow.Insert(-4121, m_objOpt);

        //    }
        //    }

        /*public readonly string TemplateFileName = @"C:\Users\Сергей\Desktop\test.docx";*/// Расположение шаблона 
        public readonly string TemplateFileName = @"Test\test.docx";// Расположение шаблона 

        private void btnWord_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox1.Text))
            {
                MessageBox.Show("Введите 'Имя'", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (string.IsNullOrEmpty(textBox2.Text))
            {
                MessageBox.Show("Введите 'Отчество'", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (string.IsNullOrEmpty(textBox3.Text))
            {
                MessageBox.Show("Введите 'Фамилию'", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (string.IsNullOrEmpty(textBox4.Text))
            {
                MessageBox.Show("Введите 'Год'", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (string.IsNullOrEmpty(textBox5.Text))
            {
                MessageBox.Show("Введите 'Год'", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }


            var lec = dataGridView2.Rows[0].Cells[0].Value.ToString(); //Ячейки
            var pra = dataGridView2.Rows[0].Cells[1].Value.ToString();
            var lab = dataGridView2.Rows[0].Cells[2].Value.ToString();

            var zach = dataGridView2.Rows[0].Cells[3].Value.ToString();
            var aykz = dataGridView2.Rows[0].Cells[4].Value.ToString();
            var kons = dataGridView2.Rows[0].Cells[5].Value.ToString();
            var praY = dataGridView2.Rows[0].Cells[6].Value.ToString();
            var summa = dataGridView2.Rows[0].Cells[11].Value.ToString();

            var Name = textBox1.Text;
            var Patronymic = textBox2.Text;
            var Surname = textBox3.Text;
            var Year = textBox4.Text;
            var Year1 = textBox5.Text;


            var wordApp = new Word.Application();//создание приложения
            wordApp.Visible = false;




            try
            {
                var fullPath = Path.GetFullPath(TemplateFileName);
               
                var wordDocument = wordApp.Documents.Open(fullPath); //открываем файл word

                ReplaceWordStub("{lec}", lec, wordDocument);// замена
                ReplaceWordStub("{pra}", pra, wordDocument);
                ReplaceWordStub("{lab}", lab, wordDocument);

                ReplaceWordStub("{zach}", zach, wordDocument);
                ReplaceWordStub("{aykz}", aykz, wordDocument);
                ReplaceWordStub("{kons}", kons, wordDocument);
                ReplaceWordStub("{praY}", praY, wordDocument);
                ReplaceWordStub("{summa}", summa, wordDocument);

                ReplaceWordStub("{Name}", Name, wordDocument);
                ReplaceWordStub("{Patronymic}", Patronymic, wordDocument);
                ReplaceWordStub("{Surname}", Surname, wordDocument);
                ReplaceWordStub("{Year}", Year, wordDocument);
                ReplaceWordStub("{Year1}", Year1, wordDocument);


                SaveFileDialog sfd = new SaveFileDialog();

                sfd.Filter = "Word Documents (*.docx)|*.docx";

                sfd.FileName = "result.docx";

                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    wordDocument.SaveAs(sfd.FileName);
                    //Export_Data_To_Word(dataGridView1, sfd.FileName);
                }




                //wordDocument.SaveAs(@"D:\result.docx");// сохраняем документ
                //wordApp.Visible = true; // отображение документа
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message/*"Произошла ошибка"*/);
            }

            //SaveFileDialog sfd = new SaveFileDialog();

            //sfd.Filter = "Word Documents (*.doc)|*.doc";

            //sfd.FileName = "export.doc";

            //if (sfd.ShowDialog() == DialogResult.OK)
            //{

            //    //ToCsV(dataGridView1, @"c:\export.xls");

            //    ToCsV(dataGridView2, sfd.FileName); // Here dataGridview1 is your grid view name 

            //}
        }

        public void ReplaceWordStub(string stubToReplace, string text, Word.Document wordDocument) // замена меток на нашу информациию
        {
            // получить содержимое документа word
            var range = wordDocument.Content;
            range.Find.ClearFormatting();// очистить поиски в документе
            range.Find.Execute(FindText: stubToReplace, ReplaceWith: text);//передача параметров (FindText-то, что хотим найти внутри документа;
                                                                           //ReplaceWith-то, чем хотим заменить)

        }









        //ненужный код
        public void Export_Data_To_Word(DataGridView DGV, string filename)
        {
            if (DGV.Rows.Count != 0)
            {
                int RowCount = DGV.Rows.Count;
                int ColumnCount = DGV.Columns.Count;
                Object[,] DataArray = new object[RowCount + 1, ColumnCount + 1];

                //add rows
                int r = 0;
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    for (r = 0; r <= RowCount - 1; r++)
                    {
                        DataArray[r, c] = DGV.Rows[r].Cells[c].Value;
                    } //end row loop
                } //end column loop

                Word.Document oDoc = new Word.Document();
                oDoc.Application.Visible = true;

                //page orintation
                oDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;


                dynamic oRange = oDoc.Content.Application.Selection.Range;
                string oTemp = "";
                for (r = 0; r <= RowCount - 1; r++)
                {
                    for (int c = 0; c <= ColumnCount - 1; c++)
                    {
                        oTemp = oTemp + DataArray[r, c] + "\t";

                    }
                }

                //table format
                oRange.Text = oTemp;

                object Separator = Word.WdTableFieldSeparator.wdSeparateByTabs;
                object ApplyBorders = true;
                object AutoFit = true;
                object AutoFitBehavior = Word.WdAutoFitBehavior.wdAutoFitContent;

                oRange.ConvertToTable(ref Separator, ref RowCount, ref ColumnCount,
                                      Type.Missing, Type.Missing, ref ApplyBorders,
                                      Type.Missing, Type.Missing, Type.Missing,
                                      Type.Missing, Type.Missing, Type.Missing,
                                      Type.Missing, ref AutoFit, ref AutoFitBehavior, Type.Missing);

                oRange.Select();

                oDoc.Application.Selection.Tables[1].Select();
                oDoc.Application.Selection.Tables[1].Rows.AllowBreakAcrossPages = 0;
                oDoc.Application.Selection.Tables[1].Rows.Alignment = 0;
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                oDoc.Application.Selection.InsertRowsAbove(1);
                oDoc.Application.Selection.Tables[1].Rows[1].Select();

                //header row style
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Bold = 1;
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Name = "Tahoma";
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Size = 14;

                //add header row manually
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    oDoc.Application.Selection.Tables[1].Cell(1, c + 1).Range.Text = DGV.Columns[c].HeaderText;
                }

                //table style 
                oDoc.Application.Selection.Tables[1].set_Style("Grid Table 4 - Accent 5");
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                oDoc.Application.Selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                //header text
                foreach (Word.Section section in oDoc.Application.ActiveDocument.Sections)
                {
                    Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);
                    headerRange.Text = "your header text";
                    headerRange.Font.Size = 16;
                    headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                }

                //save the file
                oDoc.SaveAs(filename);

                //NASSIM LOUCHANI
            }
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();

            sfd.Filter = "Word Documents (*.docx)|*.docx";

            sfd.FileName = "export.docx";

            if (sfd.ShowDialog() == DialogResult.OK)
            {

                Export_Data_To_Word(dataGridView1, sfd.FileName);
            }
        }
    }
}
