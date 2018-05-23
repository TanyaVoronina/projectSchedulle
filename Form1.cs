using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reflection;
using ExcelObj = Microsoft.Office.Interop.Excel;

namespace ProjectSchedulling
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            comboBox1.Items.AddRange(new string[] { "1 неделя", "2 неделя" });
            comboBox2.Items.AddRange(new string[] { "ЭВМ-1Н", "ЭВМ-1.2П" });
            comboBox3.Items.AddRange(new string[] { "Параллельное программирование", "Системы Исскуственного интеллекта", "Отказоустойчивые системы", "Разработка ПО"});
            comboBox4.Items.AddRange(new string[] { "Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота"});
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            //Задаем расширение имени файла по умолчанию.
            ofd.DefaultExt = "*.xls;*.xlsx";
            //Задаем строку фильтра имен файлов, которая определяет
            //варианты, доступные в поле "Файлы типа" диалогового
            //окна.
            ofd.Filter = "Excel Sheet(*.xlsx)|*.xlsx";
            //Задаем заголовок диалогового окна.
            ofd.Title = "Выберите документ для загрузки данных";
            ExcelObj.Application app = new ExcelObj.Application();
            ExcelObj.Workbook workbook;
            ExcelObj.Worksheet NwSheet;
            ExcelObj.Range ShtRange;
            DataTable dt = new DataTable();
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = ofd.FileName;

                workbook = app.Workbooks.Open(ofd.FileName, Missing.Value,
                Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                Missing.Value);

                //Устанавливаем номер листа из котрого будут извлекаться данные
                //Листы нумеруются от 1
                NwSheet = (ExcelObj.Worksheet)workbook.Sheets.get_Item(1);
                ShtRange = NwSheet.UsedRange;
                for (int Cnum = 1; Cnum <= ShtRange.Columns.Count; Cnum++)
                {
                    dt.Columns.Add(
                       new DataColumn((ShtRange.Cells[1, Cnum] as ExcelObj.Range).Value2.ToString()));
                }
                dt.AcceptChanges();

                string[] columnNames = new String[dt.Columns.Count];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    columnNames[0] = dt.Columns[i].ColumnName;
                }

                for (int Rnum = 2; Rnum <= ShtRange.Rows.Count; Rnum++)
                {
                    DataRow dr = dt.NewRow();
                    for (int Cnum = 1; Cnum <= ShtRange.Columns.Count; Cnum++)
                    {
                        if ((ShtRange.Cells[Rnum, Cnum] as ExcelObj.Range).Value2 != null)
                        {
                            dr[Cnum - 1] =
                (ShtRange.Cells[Rnum, Cnum] as ExcelObj.Range).Value2.ToString();
                        }
                    }
                    dt.Rows.Add(dr);
                    dt.AcceptChanges();
                }


                //string[,] subjects = new string[dt.Rows.Count, dt.Columns.Count];

                //for (int i = 0; i < dt.Rows.Count; i++)
                //{
                //    DataRow row = dt.Rows[i];
                //    for (int j = 0; j < dt.Columns.Count; j++)
                //    {
                //        subjects[i, j] = row[j].ToString();
                //    }
                //}

                dataGridView1.DataSource = dt;
                app.Quit();
            }
            else
                Application.Exit();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //OpenFileDialog ofd = new OpenFileDialog();
            //ofd.DefaultExt = "*.xls;*.xlsx";
            ////ofd.Filter = "Excel 2003(*.xls)|*.xls|Excel 2007(*.xlsx)|*.xlsx";
            //ofd.Title = "Выберите документ для загрузки данных";

            //if (ofd.ShowDialog() == DialogResult.OK)
            //{
            //    textBox1.Text = ofd.FileName;

            //    String constr = "Provider=Microsoft.JET.OLEDB.4.0;Data Source=" + ofd.FileName + ";Extended Properties='Excel 8.0; HDR=Yes;IMEX=1;'";

            //    System.Data.OleDb.OleDbConnection con =
            //        new System.Data.OleDb.OleDbConnection(constr);
            //    con.Open();

            //    DataSet ds = new DataSet();
            //    DataTable schemaTable = con.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables,
            //        new object[] { null, null, null, "TABLE" });

            //    string sheet1 = (string)schemaTable.Rows[0].ItemArray[2];
            //    string select = String.Format("SELECT * FROM [{0}]", sheet1);

            //    System.Data.OleDb.OleDbDataAdapter ad =
            //        new System.Data.OleDb.OleDbDataAdapter(select, con);

            //    ad.Fill(ds);

            //    DataTable tb = ds.Tables[0];
            //    con.Close();
            //    dataGridView1.DataSource = tb;
            //    con.Close();
            //}
            //else
            //{
            //    MessageBox.Show("Вы не выбрали файл для открытия",
            //            "Загрузка данных...", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
        }

        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            //Задаем расширение имени файла по умолчанию.
            ofd.DefaultExt = "*.xls;*.xlsx";
            //Задаем строку фильтра имен файлов, которая определяет
            //варианты, доступные в поле "Файлы типа" диалогового
            //окна.
            ofd.Filter = "Excel Sheet(*.xlsx)|*.xlsx";
            //Задаем заголовок диалогового окна.
            ofd.Title = "Выберите документ для загрузки данных";
            ExcelObj.Application app = new ExcelObj.Application();
            ExcelObj.Workbook workbook;
            ExcelObj.Worksheet NwSheet;
            ExcelObj.Range ShtRange;
            DataTable dt = new DataTable();

            dataGridView1.AutoSizeRowsMode =
        DataGridViewAutoSizeRowsMode.AllCells;

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = ofd.FileName;

                workbook = app.Workbooks.Open(ofd.FileName, Missing.Value,
                Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                Missing.Value);

                //Устанавливаем номер листа из котрого будут извлекаться данные
                //Листы нумеруются от 1
                NwSheet = (ExcelObj.Worksheet)workbook.Sheets.get_Item(1);
                ShtRange = NwSheet.UsedRange;
                for (int Cnum = 1; Cnum <= ShtRange.Columns.Count; Cnum++)
                {
                    dt.Columns.Add(
                       new DataColumn((ShtRange.Cells[1, Cnum] as ExcelObj.Range).Value2.ToString()));
                }
                dt.AcceptChanges();

                string[] columnNames = new String[dt.Columns.Count];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    columnNames[0] = dt.Columns[i].ColumnName;
                }

                for (int Rnum = 2; Rnum <= ShtRange.Rows.Count; Rnum++)
                {
                    DataRow dr = dt.NewRow();
                    for (int Cnum = 1; Cnum <= ShtRange.Columns.Count; Cnum++)
                    {
                        if ((ShtRange.Cells[Rnum, Cnum] as ExcelObj.Range).Value2 != null)
                        {
                            dr[Cnum - 1] =
                (ShtRange.Cells[Rnum, Cnum] as ExcelObj.Range).Value2.ToString();
                        }
                    }
                    dt.Rows.Add(dr);
                    dt.AcceptChanges();
                }

                //пары -> массив
                string[,] subjects = new string[dataGridView1.Rows.Count, dataGridView1.ColumnCount];

                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView1.ColumnCount; j++)
                    {
                        subjects[i, j] = Convert.ToString(dataGridView1.Rows[i].Cells[j].Value);
                    }
                }

                //извлечение критериев
                int indexWeek;
                int indexGroup;
                int indexSub;
                int indexDay;
                string sub;

                indexWeek = comboBox1.SelectedIndex;
                indexGroup = comboBox2.SelectedIndex;
                indexSub = comboBox3.SelectedIndex;
                indexDay = comboBox4.SelectedIndex;

                //если выбрана первая неделя
                if (indexWeek == 0)
                {
                    //вставляем в первый пропуск, получаем из виджета предмет и для какой группы, вставка в массив
                    if(indexDay == 0)
                    {
                        if (indexGroup == 0)
                        { 
                            for(int i = 0; i < 5; i++)
                            {
                                for (int j = 0; j < 1; j++)
                                {
                                    if ((subjects[i, j].Equals("")))
                                    {
                                        sub = comboBox3.Text;
                                        subjects[i + 1, j] = sub;
                                    }
                                    else
                                    { }
                                }
                            }
                        }
                        else 
                        { 

                        }
                    }
                    //если понедельник - перебираем первые 6 элементов массива, вставка
                    else if(indexDay == 1)
                    {
                        if (indexGroup == 0)
                        { 

                        }
                        else
                        {

                        }
                    }
                    //если вторник вторые - перебираем первые 6 элементов массива, вставка и т.д.
                    else if(indexDay == 2)
                    {
                        if (indexGroup == 0)
                        { 

                        }
                        else
                        { 

                        }
                    }

                    else if(indexDay == 3)
                    {
                        if (indexGroup == 0)
                        {

                        }
                        else
                        { 

                        }
                    }

                    else if(indexDay == 4)
                    {
                        if (indexGroup == 0)
                        {

                        }
                        else
                        {

                        }
                    }

                    else if (indexDay == 5)
                    {
                        if (indexGroup == 0)
                        {

                        }
                        else
                        {

                        }
                    }

                    
                    for (int i = 0; i < subjects.GetLength(0); i++)
                    {
                        for (int j = 0; j < subjects.GetLength(1); j++)
                        {
                            //пишем значения из массива в ячейки контролла
                            dataGridView1.Rows[i].Cells[j].Value = subjects[i, j];
                        }
                    }
                }


                //если выбрана вторая неделя
                else
                { 

                }

            }

        }

    }
}

