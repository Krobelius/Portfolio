using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows;
using System.Windows.Controls;
using System.Data;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace WEReplace1._0
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        bool check_box = false;
        public DataTable excel_data;
        public string[] array_path;
        public string def_path;
        public MainWindow()
        {
            InitializeComponent();
        }
        private void Window_Initialized(object sender, EventArgs e)
        {
            
            Props pr = new Props();
            pr.ReadXml();
            def_path = pr.Fields.path_value;
            if(def_path == null)
            {
                def_path = Environment.CurrentDirectory;
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
        }

        public void Button_Click_1(object sender, RoutedEventArgs e)
        {
            FilesWork fw = new FilesWork();
            try
            {
                string xlsx_file = fw.files_connect(true);
                if (xlsx_file != "null")
                {
                    files_box.Items.Add(xlsx_file);
                    excel_data = fw.ConvExDt(xlsx_file);
                    files_box.Items.Add("Файл выбран успешно!");
                }
            }
            catch (Exception)
            {
                System.Windows.Forms.MessageBox.Show("Произошла ошибка!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            FilesWork fw = new FilesWork();
            try
            {
                string docx_file = fw.files_connect(false);
                if (docx_file != "null")
                {
                    array_path = docx_file.Split('|');
                    foreach (string i in array_path)
                    {
                        files_box.Items.Add(i);
                        files_box.Items.Add("Файл выбран успешно!");
                    }
                }
            }
            catch (Exception)
            {
                System.Windows.Forms.MessageBox.Show("Произошла ошибка!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        public void ListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        public void Button_Click_3(object sender, RoutedEventArgs e)
        {
            FilesWork fw = new FilesWork();
            try
            {
                fw.OpenAndReplace(array_path, def_path, excel_data,check_box);
                files_box.Items.Add("Данные заменены успешно!");
            }
            catch (Exception)
            {
                System.Windows.Forms.MessageBox.Show("Произошла ошибка!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {
            check_box = true;
        }

        private void Button_Click_5(object sender, RoutedEventArgs e)
        {

        }

        private void Button_Click_6(object sender, RoutedEventArgs e)
        {
            Help.ShowHelp(null, "Справка.chm");
        }

        private void Button_Click_Path(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Props pr = new Props();
                pr.Fields.path_value = fbd.SelectedPath;
                def_path = fbd.SelectedPath;
                pr.WriteXml();
            }
        }

        private void Window_StateChanged(object sender, EventArgs e)
        {
        }

        private void Button_Click_4(object sender, RoutedEventArgs e)
        {
            Environment.Exit(0);
        }
    }
}
