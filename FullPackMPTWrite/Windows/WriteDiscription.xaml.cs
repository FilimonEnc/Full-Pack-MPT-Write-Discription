using System.Windows;
using System.Windows.Controls;

using MahApps.Metro.Controls;
using Microsoft.Win32;
using WK.Libraries.BetterFolderBrowserNS;

namespace FullPackMPTWrite.Windows
{
    /// <summary>
    /// Логика взаимодействия для WriteDiscription.xaml
    /// </summary>
    public partial class WriteDiscription : MetroWindow
    {
        public WriteDiscription()
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
        private void buttonFillDiscription_Click(object sender, RoutedEventArgs e)
        {
            OpenFile(pathDiscription);
        }

        private void buttonFillFolder_Click(object sender, RoutedEventArgs e)
        {
            OpenFolder(pathFolder);
        }

        private void buttonStart_Click(object sender, RoutedEventArgs e)
        {
            ScriptsWriteDiscription.Excel ex = new ScriptsWriteDiscription.Excel(pathDiscription.Content.ToString(), 0);

            ex.Work(pathFolder.Content.ToString());
            ex.Close();
        }

        private void buttonStartPractic_Copy_Click(object sender, RoutedEventArgs e)
        {
            ScriptsWriteDiscription.Excel ex = new ScriptsWriteDiscription.Excel(pathDiscriptionPractic.Content.ToString(), 1);

            ex.WorkPractic(pathFolder.Content.ToString());
            ex.Close();
        }

        private void buttonFillDiscriptionPractic_Click(object sender, RoutedEventArgs e)
        {
            OpenFile(pathDiscriptionPractic);
        }
    }
}
