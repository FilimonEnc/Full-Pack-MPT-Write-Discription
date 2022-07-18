using MahApps.Metro.Controls;
using Microsoft.Win32;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using FullPackMPTWrite.ScriptsCreateDiscription;
using WK.Libraries.BetterFolderBrowserNS;

namespace FullPackMPTWrite.Windows
{

    public partial class CreateDiscription : MetroWindow
    {

        SaveFileDialog saveFileDialog = new SaveFileDialog();
        public CreateDiscription()
        {
            InitializeComponent();
        }
        public void OpenFile(Label label)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel files(*.xlsx)|*.xlsx|All files(*.*)|*.*";
            if (openFileDialog.ShowDialog() == false)
                return;
            string filename = openFileDialog.FileName.ToString();
            label.Content = filename;
        }
        public void OpenFolder(Label label)
        {

            BetterFolderBrowser folderBrowserDialog = new BetterFolderBrowser();
            if (folderBrowserDialog.ShowDialog() == System.Windows.Forms.DialogResult.Cancel)
                return;
            string filename = folderBrowserDialog.SelectedPath;
            label.Content = filename;
        }
        private void buttonFillBlank_Click(object sender, RoutedEventArgs e)
        {
            OpenFile(pathBlank);
        }

        private void buttonFillSheet_Click(object sender, RoutedEventArgs e)
        {
            OpenFile(pathSheet);
        }

        private void buttonFillFolder_Click_1(object sender, RoutedEventArgs e)
        {
            OpenFolder(pathSave);
        }

        private void buttonStart_Click(object sender, RoutedEventArgs e)
        {
            bool chek = true;
            int i = 1;
            while (chek)
            {
                try
                {
                    
                    Data data = new Data();
                    Excel ex = new Excel(pathSheet.Content.ToString(), i);
                    Data read = ex.ReadRange();
                    string inicial = ex.SendInfo(read);
                    ex.Close();
                    Excel ex1 = new Excel(pathBlank.Content.ToString(), 12);
                    ex1.WriteRange(read);

                    ex1.SaveAs(@"" + pathSave.Content.ToString() + "\\" + "Ведомость учета часов " + inicial + ".xlsx");
                    ex1.Close();
                    i++;
                }
                catch
                {
                    chek = false;
                }
            }

            MessageBox.Show("файл создан! выберите следующую нагрузку");
        }

        private void buttonStartally_Click(object sender, RoutedEventArgs e)
        {
            List<string> pathfolder = GetFolders(pathFolderAll.Content.ToString());


            foreach (string value in pathfolder)
            {
                int i = 1;
                bool chek = true;
                while (chek)
                {
                    try
                    {
                        
                        Data data = new Data();
                        Excel ex = new Excel(value, i);
                        Data read = ex.ReadRange();
                        string inicial = ex.SendInfo(read);
                        ex.Close();
                        Excel ex1 = new Excel(pathBlank.Content.ToString(), 12);
                        ex1.WriteRange(read);

                        ex1.SaveAs(@"" + pathSave.Content.ToString() + "\\" + "Ведомость учета часов " + inicial + ".xlsx");
                        ex1.Close();
                        i++;
                    }
                    catch
                    {
                        chek = false;
                    }
                }

            }
            MessageBox.Show("Все файлы созданы");
        }

        private void buttonFillfooldernagruzki_Click(object sender, RoutedEventArgs e)
        {
            OpenFolder(pathFolderAll);
        }
        public List<string> GetFolders(string _folder)
        {
            string folders = _folder;
            DirectoryInfo dir = new DirectoryInfo(folders);
            FileInfo[] files = dir.GetFiles();
            List<string> pathFileName = new List<string>();
            foreach (var value in files)
            {
                pathFileName.Add(_folder + "\\" + value.ToString());
            }

            return pathFileName;
        }
    }
}
