﻿using System;
using System.IO;
using NPOI.XSSF.UserModel;
using System.Data;
using System.Linq;
using Microsoft.Office.Interop;
using System.Windows.Forms;
using System.Text;
using System.Threading.Tasks;

namespace WEReplace1._0
{
    class FilesWork
    {
        public string files_connect(bool file)
        {
            OpenFileDialog fd = new OpenFileDialog { Multiselect = true, Title = "Выберите нужные Word/Excel файлы", Filter = "Image Files (doc,docx)| *.doc;*.docx" };
            if (file)
            {
                fd.Multiselect = false;
                fd.Filter = "Image Files (xlsx,xls)|*.xlsx;*.xls";
            }
            if (fd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                if (file)
                    return fd.FileName;
                else
                    return (String.Join("|", fd.FileNames));
            }
            return ("null");
        }
        public System.Data.DataTable ConvExDt(string path)
        {
            XSSFWorkbook wb = null;
            XSSFSheet sh = null;
            string sheet_name = null ;
            try
            {
                using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read))
                {
                    wb = new XSSFWorkbook(fs);
                    sheet_name = wb.GetSheetAt(0).SheetName;
                }
                sh = (XSSFSheet)wb.GetSheet(sheet_name);
            }
            catch (Exception e)
            {
                MessageBox.Show("Error!", "Проблема с открытием файла. Проверьте, что файл Excel выбран правильно и не поврежден" + e.Message, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Rows.Clear();
            dt.Columns.Clear();
            int i = 0;
            //тут необходимо проверить DataTable на наличие нужного кол-ва столбцов, чтобы при добавлении данных не выкидывало ошибку
            while(sh.GetRow(i) != null)
            {
                if(dt.Columns.Count < sh.GetRow(i).Cells.Count)
                {
                    for(int j = 0;j < sh.GetRow(i).Cells.Count;j++)
                    { 
                        dt.Columns.Add("", typeof(string));
                    }
                }
                //и добавляем строку
                dt.Rows.Add();
                //заполняем данными из excel, сравнивая типы через Case. Пришлось добавить отдельное сравнение для дата-типа, в CellType его нет.
                for(int j = 0;j < sh.GetRow(i).Cells.Count;j++)
                {
                    var cell = sh.GetRow(i).GetCell(j);
                    if(DateTime.TryParse(cell.ToString(),out DateTime datevalue))
                    {
                        dt.Rows[i][j] = Convert.ToString(cell.DateCellValue.Date.Date);
                        continue;
                    }

                    if(cell != null)
                    {
                        switch(cell.CellType)
                        {
                            case NPOI.SS.UserModel.CellType.Numeric:
                                dt.Rows[i][j] = sh.GetRow(i).GetCell(j).NumericCellValue;
                                break;

                            case NPOI.SS.UserModel.CellType.String:
                                dt.Rows[i][j] = sh.GetRow(i).GetCell(j).StringCellValue;
                                break;
                        }
                    }
                }
                i++;
            }
            return dt;
        }
        public int CheckingRows(object[] item_arr)
        {
            int summ_rows = 0;
                foreach(var k in item_arr)
                {
                    if(k.ToString() != "")
                    {
                        summ_rows++;
                    }
                }
            return summ_rows;
        }
        //дальше идут методы, необходимые для работы с word-файлами в Interop. Я не нашел в NPOI такого же удобного функционала для Word, как для Excel, с этим и связан переход на Interop
        //обычно я бы предпочел не работать с ним, т.к Interop требует установленного Office на ПК и может вызывать ошибки, но для C# не так много удобных библиотек для работы с Office
        Microsoft.Office.Interop.Word.Application wordApp;
        public void OpenAndReplace(string [] Fnames,string def_path,System.Data.DataTable dt,bool check)
        {
            for (int n = 0;n<Fnames.Length;n++)
            {
                try
                {

                    wordApp = new Microsoft.Office.Interop.Word.Application();
                    wordApp.Documents.Open(Fnames[n], Type.Missing, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                    var tbls = wordApp.ActiveDocument.Tables.Count;
                    for (int j = 1; j <= tbls; j++)
                    {
                        Microsoft.Office.Interop.Word.Table tbl = wordApp.ActiveDocument.Tables[j];
                        foreach (System.Data.DataRow row in dt.Rows)
                        {
                            for (int t = 1; t <= tbl.Columns.Count; t++)
                            {
                                string fixed_text = tbl.Cell(1, t).Range.Text;
                                fixed_text = fixed_text.Replace("\r", "");
                                fixed_text = fixed_text.Replace("\a", "");
                                if (row.ItemArray[0].ToString() == fixed_text)
                                {
                                    if (tbl.Rows.Count < CheckingRows(row.ItemArray))
                                    {
                                        for (int rw = CheckingRows(row.ItemArray) - tbl.Rows.Count; rw > 0; rw--)
                                        {
                                            tbl.Rows.Add();
                                        }
                                    }
                                    for (int k = 0; k < CheckingRows(row.ItemArray); k++)
                                    {
                                        DateTime date;
                                        if(DateTime.TryParse(row.ItemArray[k].ToString(),out date))
                                        {
                                            var date_replace = Convert.ToDateTime(row.ItemArray[k]).ToLongDateString();
                                            tbl.Cell(k + 1, t).Range.Text = date_replace.ToString();
                                            tbl.Cell(k + 1, t).Range.Font.Bold = 0;
                                            tbl.Cell(k + 1, t).Range.Font.AllCaps = 0;
                                        }
                                        else
                                        {
                                            tbl.Cell(k + 1, t).Range.Text = row.ItemArray[k].ToString();
                                            tbl.Cell(k + 1, t).Range.Font.Bold = 0;
                                            tbl.Cell(k + 1, t).Range.Font.AllCaps = 0;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show("Проблема с открытием файла. Проверьте, что файл Word выбран правильно и не поврежден" + e.Message, "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    wordApp.ActiveDocument.Close();
                }
                DateTime datevalue;
                try
                {
                    var doc = wordApp.ActiveDocument;
                    foreach(Microsoft.Office.Interop.Word.Range range in doc.StoryRanges)
                    {
                        for(int k = 1;k<=range.Paragraphs.Count;k++)
                        {
                            if(range.Paragraphs[k].Range.Tables.Count != 0)
                            {
                                continue;
                            }
                            else
                            {
                                var range_not_table = range.Paragraphs[k].Range;
                                try
                                {
                                    for (int i = 0; i < dt.Rows.Count; i++)
                                    {
                                        for (int j = 1; j < dt.Rows[i].ItemArray.Length; j++)
                                        {
                                            string value;
                                            if (check)
                                            {
                                                value = dt.Rows[i].ItemArray[1].ToString();
                                            }
                                            else
                                            {
                                                value = dt.Rows[i].ItemArray[j].ToString();
                                            }
                                            if (DateTime.TryParse(value.ToString(), out datevalue) == false)
                                            {
                                                range_not_table.Find.Replacement.Text = value;
                                                range_not_table.Find.Replacement.Font.Bold = 0;
                                                range.Find.Replacement.Font.AllCaps = 0;
                                                range_not_table.Find.Execute(dt.Rows[i].ItemArray[0].ToString(), true, true, false, Type.Missing, Type.Missing, true, Type.Missing, true, Type.Missing, Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll, false, false, false, false);
                                            }
                                            else
                                            {
                                                range_not_table.Find.Replacement.Text = Convert.ToDateTime(value).ToLongDateString();
                                                range_not_table.Find.Execute(dt.Rows[i].ItemArray[0].ToString(), true, true, false, Type.Missing, Type.Missing, true, Type.Missing, true, Type.Missing, Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll, false, false, false, false);
                                            }
                                        }
                                    }
                                    wordApp.ActiveDocument.SaveAs2(def_path + "\\" + Path.GetFileNameWithoutExtension(Fnames[n]) + " Готовый" + ".docx");
                                }
                                catch (Exception e)
                                {
                                    MessageBox.Show("Произошла ошибка при попытке записи данных в файл" + e.Message, "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    wordApp.ActiveDocument.Close();
                                }
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show("Произошла ошибка!" +e.Message, "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    wordApp.ActiveDocument.Close();
                }
                wordApp.ActiveDocument.Close();
            }
        }
    }
}
