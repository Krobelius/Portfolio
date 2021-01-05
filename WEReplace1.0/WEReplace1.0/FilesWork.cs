using System;
using System.IO;
using NPOI.XSSF.UserModel;
using System.Data;
using Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using System.Linq;
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
        //дальше идут методы, необходимые для работы с word-файлами в Interop. Я не нашел в NPOI такого же удобного функционала для Word, как для Excel, с этим и связан переход на Interop
        //обычно я бы предпочел не работать с ним, т.к Interop требует установленного Office на ПК и может вызывать ошибки, но для C# не так много удобных библиотек для работы с Office
        Microsoft.Office.Interop.Word.Application wordApp;
        public void Table_Filler(string[] Fnames, string def_path, System.Data.DataTable dt)
        {
            for (int n = 0; n < Fnames.Length; n++)
            {
                try
                {
                    wordApp = new Microsoft.Office.Interop.Word.Application();
                    wordApp.Documents.Open(Fnames[n], Type.Missing, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                    Microsoft.Office.Interop.Word.Table tbl = wordApp.ActiveDocument.Tables[1];
                    foreach (System.Data.DataRow row in dt.Rows)
                    {
                        for (int t = 1; t <= tbl.Columns.Count; t++)
                        {
                            string fixed_text = tbl.Cell(1, t).Range.Text;
                            fixed_text = fixed_text.Replace("\r", "");
                            fixed_text = fixed_text.Replace("\a", "");
                            if (row.ItemArray[0].ToString() == fixed_text)
                            {
                                for (int k = 0; k < row.ItemArray.Length; k++)
                                {
                                    tbl.Cell(k + 1, t).Range.Text = row.ItemArray[k].ToString();
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
                try
                {
                    wordApp.ActiveDocument.SaveAs2(def_path + "ND" + (n + 1) + ".docx");
                }
                catch (Exception e)
                {
                    MessageBox.Show("Произошла ошибка при попытке записи данных в файл" + e.Message, "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    wordApp.ActiveDocument.Close();
                }
                wordApp.ActiveDocument.Close();
            }
        }
        public void OpenAndReplace(string [] Fnames,string def_path,System.Data.DataTable dt,bool check)
        {
            for(int n = 0;n<Fnames.Length;n++)
            {
                DateTime datevalue;
                try
                {
                    wordApp = new Microsoft.Office.Interop.Word.Application();
                    Object template = Fnames[n];
                    Object newTemplate = true;
                    Object documentType = Microsoft.Office.Interop.Word.WdNewDocumentType.wdNewBlankDocument;
                    Object visible = true;
                    wordApp.Documents.Open(Fnames[n], Type.Missing, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                }
                catch (Exception e)
                {
                    MessageBox.Show("Произошла ошибка!" +e.Message, "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    wordApp.ActiveDocument.Close();
                }
                try
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                      for (int j = 0; j < dt.Rows[i].ItemArray.Length; j++)
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
                    if (DateTime.TryParse(value.ToString(),out datevalue) == false)
                            {
                                ReplaceText(dt.Rows[i].ItemArray[0].ToString(), value);
                            }
                    else
                            {
                                ReplaceText(dt.Rows[i].ItemArray[0].ToString(), Convert.ToDateTime(value).ToLongDateString());
                            }
                        }
                    }
                    wordApp.ActiveDocument.SaveAs2(def_path + "ND" + (n+1) + ".docx");
                }
                catch (Exception e)
                {
                    MessageBox.Show("Произошла ошибка при попытке записи данных в файл" + e.Message, "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    wordApp.ActiveDocument.Close();
                }
                wordApp.ActiveDocument.Close();
            }
        }
        public void ReplaceText(string word, string repl)
        {
            Object unit = Microsoft.Office.Interop.Word.WdUnits.
            wdStory;
            Object extend = Microsoft.Office.Interop.Word.
            WdMovementType.wdMove;
            wordApp.Selection.HomeKey(ref unit, ref extend);
            Microsoft.Office.Interop.Word.Find fnd = wordApp.Selection.
            Find;
            fnd.ClearFormatting();
            fnd.Text = word;
            fnd.Replacement.ClearFormatting();
            fnd.Replacement.Text = repl;
            ExecuteReplace(fnd);
        }
        private Boolean ExecuteReplace(Microsoft.Office.Interop.Word.
Find find)
        {
            return ExecuteReplace(find, Microsoft.Office.Interop.Word.
            WdReplace.wdReplaceAll);

        }
        private Boolean ExecuteReplace(Microsoft.Office.Interop.Word.
Find find, Object replaceOption)
        {
            Object findText = Type.Missing;
            Object matchCase = Type.Missing;
            Object matchWholeWord = Type.Missing;
            Object matchWildcards = Type.Missing;
            Object matchSoundsLike = Type.Missing;
            Object matchAllWordForms = Type.Missing;
            Object forward = Type.Missing;
            Object wrap = Type.Missing;
            Object format = Type.Missing;
            Object replaceWith = Type.Missing;
            Object replace = replaceOption;
            Object matchKashida = Type.Missing;
            Object matchDiacritics = Type.Missing;
            Object matchAlefHamza = Type.Missing;
            Object matchControl = Type.Missing;
            return find.Execute(ref findText, ref matchCase,
            ref matchWholeWord, ref matchWildcards, ref matchSoundsLike,
            ref matchAllWordForms, ref forward, ref wrap, ref format,
            ref replaceWith, ref replace, ref matchKashida,
            ref matchDiacritics, ref matchAlefHamza, ref matchControl);
        }
    }
}
