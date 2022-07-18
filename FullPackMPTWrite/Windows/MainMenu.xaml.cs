using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using MahApps.Metro.Controls;

namespace FullPackMPTWrite.Windows
{
    /// <summary>
    /// Логика взаимодействия для MainMenu.xaml
    /// </summary>
    public partial class MainMenu : MetroWindow
    {
        public MainMenu()
        {
            InitializeComponent();
        }

        private void buttonCreate_Click(object sender, RoutedEventArgs e)
        {
            CreateDiscription n = new CreateDiscription();
            n.Show();
        }

        private void buttonWrite_Click(object sender, RoutedEventArgs e)
        {
            WriteDiscription n = new WriteDiscription();
            n.Show();
        }

        private void buttonWrite_Copy_Click(object sender, RoutedEventArgs e)
        {
            SQLiteConnection con = new SQLiteConnection("Data Source=./MPTList.db");
            SQLiteDataAdapter da, da1;
            SQLiteCommand cmd;
            DataSet ds;
            DataTable dt;
            SQLiteDataReader reader;
            string error = "";

            cmd = new SQLiteCommand("DELETE FROM Sheet; DELETE FROM SheetGroup; DELETE FROM Groupes; DELETE FROM Subject; DELETE FROM Teacher;", con);
            if (!CommectionOpen(ref con, ref error))
            {
                Debug.WriteLine("MYDEBUG: " + error);
                MessageBox.Show("Подключение к базе данных отсутсвует");
                //return;
            }
            reader = cmd.ExecuteReader();
            reader.Close();
            con.Close();
        }
        bool CommectionOpen(ref SQLiteConnection conn, ref string error)
        {
            try
            {
                conn.Open();
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }

            return true;
        }
    }
}
