using System;
using System.Collections.Generic;
using System.Data;
//using System.Data.SqlClient;
using System.Data.SQLite;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;
using OfficeOpenXml;

namespace FullPackMPTWrite.ScriptsCreateDiscription
{
    struct Data
    {
        public object[,] data;
        public object[,] dataGroup;
        public object[,] data1Semestr;
        public object[,] data2Semestr;
        public object fio;
        public int end;
        public string inicial;
    }


    class Excel
    {
        string connectionString = "Data Source=./MPTList.db";
        SQLiteConnection con = new SQLiteConnection("Data Source=./MPTList.db");
        SQLiteDataAdapter da, da1;
        SQLiteCommand cmd;
        DataSet ds;
        DataTable dt;


        public int g, v;
        public int numendi;
        string path = "";
        //_Application excel = new _Excel.Application();
        ExcelPackage excel;
        //Workbook wb;
        ExcelWorkbook wb;
        //Worksheet ws;
        ExcelWorksheet ws;

        FileStream fs;

        public Excel(string path, int Sheet)
        {

            // If you use EPPlus in a noncommercial context
            // according to the Polyform Noncommercial license:
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            this.path = path;
            fs = new FileStream(path, FileMode.Open);
            excel = new ExcelPackage(fs);
            //wb = excel.Workbooks.Open(path);
            wb = excel.Workbook;
            //ws = wb.Worksheets[Sheet];
            ws = wb.Worksheets[Sheet];
        }


        public int FindCell()
        {
            int i = 1;
            int j = 1;

            while (true)
            {
                if (Convert.ToString(ws.Cells[i, j].Value) == Convert.ToString("               В С Е Г О"))
                {
                    //MessageBox.Show("нашел " + i);
                    //9
                    g = i;
                    return g;
                }
                else
                {
                    i++;
                    // MessageBox.Show("не нашел ");
                }

            }

        }
        public int FindCell2()
        {
            int o = 1;
            int k = 1;
            while (true)
            {

                if (Convert.ToString(ws.Cells[o, k].Value) == Convert.ToString("1"))
                {
                    v = o;
                    return v;
                }
                else
                {
                    o++;

                }

            }
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
                MessageBox.Show(error);
                return false;
            }

            return true;
        }

        public Data ReadRange()
        {
            object fioObject = ws.Cells[7, 18].Value;
            string fio = fioObject.ToString();
            string[] words = fio.Split(new char[] { ' ' });
            string firstName = words[1];
            string secondName = words[0];
            string lastName = words[2];
            firstName = firstName.Trim();
            secondName = secondName.Trim();
            lastName = lastName.Trim();
            SQLiteDataReader reader = null;
            string tempinicial = "";
            string error = "";


            //Init please
            /*

            da = new SqlDataAdapter("SELECT Abbreviation FROM Teacher WHERE FirstName  ='" + firstName  + "' " + "AND [SurName] ='" + secondName  + "'", con);
            dt = new DataTable();
            string tempinicial ="";
            try
            {
                con.Open();
                da.Fill(dt);
                tempinicial = dt.Rows[0].ToString();
                con.Close();
            }
            catch(Exception ex)
            { 
                MessageBox.Show(@"Такого преподавателя нет в базе данных. Пожулайста заполните базу данных. Возможно программа сейчас сломается");

            }
            */
            int starti;
            int starty = 2;
            int endi;
            int endy = 2;
            int startyGroup = 3;
            int endyGroup = 3;
            int starty1Semestr = 14;
            int endy1Semestr = 14;
            int starty2Semestr = 29;
            int endy2Semestr = 29;


            FindCell();
            //  FindCell2();
            endi = g - 1;
            starti = 14;
            numendi = endi - starti + 9;
            ExcelRange range = ws.Cells[starti, starty, endi, endy];
            ExcelRange rangeGroup = ws.Cells[starti, startyGroup, endi, endyGroup];
            ExcelRange range1Semestr = ws.Cells[starti, starty1Semestr, endi, endy1Semestr];
            ExcelRange range2Semestr = ws.Cells[starti, starty2Semestr, endi, endy2Semestr];
            ExcelRange rangefio = ws.Cells[7, 18];
            object holder = range.Value;
            object holderGroup = rangeGroup.Value;
            object holder1Semestr = range1Semestr.Value;
            object holder2Semestr = range2Semestr.Value;
            object holderfio = rangefio.Value;

            //string[,] returnstring = new string[endi - starti + 1, endy - starty ];
            //тут вызывем запись в бд
            WriteDB(holderGroup, holder, holderfio);

            cmd = new SQLiteCommand("select [Teacher].[Abbreviation] from [Teacher] where [Teacher].[FirstName] = @FirstName and [Teacher].[SurName] = @LastName", con);

            if (!CommectionOpen(ref con, ref error))
            {
                Debug.WriteLine("MYDEBUG: " + error);
                MessageBox.Show("Подключение к базе данных отсутсвует");
                return new Data();

            }

            cmd.Parameters.AddWithValue("@FirstName", firstName);
            cmd.Parameters.AddWithValue("@LastName", secondName);

            reader = cmd.ExecuteReader();

            if (reader.HasRows)
            {
                if (reader.Read())
                {
                    tempinicial = reader.GetString(0);
                }
            }

            reader.Close();

            con.Close();

            Debug.WriteLine("MYDEBUG: " + tempinicial);

            Data data = new Data();

            data.data = (object[,])holder;
            data.dataGroup = (object[,])holderGroup;
            data.data1Semestr = (object[,])holder1Semestr;
            data.data2Semestr = (object[,])holder2Semestr;
            data.end = numendi;
            data.fio = holderfio;
            data.inicial = tempinicial;
            Debug.WriteLine("DEBUG: " + range.Value);

            Debug.WriteLine("DEBUG: " + ((object[,])holder).Length);

            Debug.WriteLine("DEBUG: " + (endi - starti + 1));
            Debug.WriteLine("DEBUG: " + (endy - starty));


            return data;

        }

        public void WriteDB(object _groups, object _subject, object _fio)
        {
            SQLiteDataReader reader = null;
            string tempinicial = "";
            string error = "";
            int idTeacher = 0;
            int idSubject = 0;
            int idGroup = 0;
            int tempSheet = 0;
            int tempGroup = 0;
            int tempSheetGroup = 0;
            string finalSubject = "";
            string fullSubject = "";
            //List<string> groupList = _groups as List<string>;
            List<string> groupList = new List<string>();

            foreach (object value in (object[,])_groups)
            {
                if (value != null)
                    groupList.Add(value.ToString());
            }


            var tempGroupList = groupList.Distinct();
            int temp = 0;
            int tempSubject = 0;
            foreach (string valueGroup in tempGroupList)
            {
                if (valueGroup != null || valueGroup != "")
                {
                    //сделать проверку нет ли уже такого значения в базе данных. и вообще работает ли это говно проверить по брек поинту
                    cmd = new SQLiteCommand(
                        "select ID_Group from [Groupes] where GroupName = @GroupNamee",
                        con);

                    if (!CommectionOpen(ref con, ref error))
                    {
                        Debug.WriteLine("MYDEBUG: " + error);
                        MessageBox.Show("Подключение к базе данных отсутсвует");
                        return;

                    }

                    cmd.Parameters.AddWithValue("@GroupNamee", valueGroup);

                    reader = cmd.ExecuteReader();

                    if (reader.HasRows)
                    {
                        if (reader.Read())
                        {
                            temp = reader.GetInt32(0);
                        }
                    }

                    reader.Close();

                    con.Close();

                    if (temp == 0)
                    {
                        cmd = new SQLiteCommand("Insert Into Groupes (GroupName) Values(@GroupName)",
                            con);

                        if (!CommectionOpen(ref con, ref error))
                        {
                            Debug.WriteLine("MYDEBUG: " + error);
                            MessageBox.Show("Подключение к базе данных отсутсвует");
                            return;

                        }

                        cmd.Parameters.AddWithValue("@GroupName", valueGroup);

                        reader = cmd.ExecuteReader();

                        reader.Close();

                        con.Close();
                    }

                    cmd = new SQLiteCommand(
                        "select ID_Group from Groupes where GroupName = @GroupNamee",
                        con);

                    if (!CommectionOpen(ref con, ref error))
                    {
                        Debug.WriteLine("MYDEBUG: " + error);
                        MessageBox.Show("Подключение к базе данных отсутсвует");
                        return;

                    }

                    cmd.Parameters.AddWithValue("@GroupNamee", valueGroup);

                    reader = cmd.ExecuteReader();

                    if (reader.HasRows)
                    {
                        if (reader.Read())
                        {
                            idGroup = reader.GetInt32(0);
                        }
                    }

                    reader.Close();

                    con.Close();
                    //запись в бд с проверкой нет ли такой хуйни
                }
            }

            //переписать как пример поовыше
            List<string> subjectList = new List<string>();
            foreach (object value in (object[,])_subject)
            {
                if (value != null)
                    subjectList.Add(value.ToString());
            }

            foreach (string valueSubject in subjectList)
            {
                if (valueSubject != null || valueSubject != "")
                {

                    string[] tempSplit = valueSubject.Split(new[] { ' ' });
                    var tempResult = tempSplit.Skip(2);

                    foreach (var valueResult in tempResult)
                    {
                        finalSubject = finalSubject + " " + valueResult;
                    }

                    finalSubject = finalSubject.Trim();

                    //запись в бд с проверкой нет ли такой хуйни
                    cmd = new SQLiteCommand(
                        "select [Subject].[ID_Subject] from [Subject] where [Subject].[Subject] = @SubjectName and [Subject].[SubjectFullName] = @FullName",
                        con);

                    if (!CommectionOpen(ref con, ref error))
                    {
                        Debug.WriteLine("MYDEBUG: " + error);
                        MessageBox.Show("Подключение к базе данных отсутсвует");
                        return;

                    }

                    cmd.Parameters.AddWithValue("@SubjectName", finalSubject);
                    cmd.Parameters.AddWithValue("@FullName", valueSubject);

                    reader = cmd.ExecuteReader();

                    if (reader.HasRows)
                    {
                        if (reader.Read())
                        {
                            tempSubject = reader.GetInt32(0);
                        }
                    }

                    reader.Close();

                    con.Close();

                    if (tempSubject == 0)
                    {
                        cmd = new SQLiteCommand(
                            "Insert Into [Subject] (Subject, SubjectFullName) Values(@SubjectName, @FullName)",
                            con);

                        if (!CommectionOpen(ref con, ref error))
                        {
                            Debug.WriteLine("MYDEBUG: " + error);
                            MessageBox.Show("Подключение к базе данных отсутсвует");
                            return;

                        }

                        cmd.Parameters.AddWithValue("@SubjectName", finalSubject);
                        cmd.Parameters.AddWithValue("@FullName", valueSubject);

                        reader = cmd.ExecuteReader();

                        reader.Close();

                        con.Close();

                    }
                    tempSubject = 0;
                    finalSubject = "";



                }
            }



            string fio = _fio.ToString();
            string[] fioo = fio.Split(' ');
            string firstName = fioo[1];
            string lastName = fioo[0];
            string middleName = fioo[2];
            char firstNameChar = firstName.FirstOrDefault();
            char middleNameChar = middleName.FirstOrDefault();
            string initial;
            initial = firstNameChar + "." + middleNameChar + ". " + lastName;

            bool chek = false;
            cmd = new SQLiteCommand(
                "SELECT [Teacher].[ID_Teacher] from [Teacher] where [Teacher].[Abbreviation] = @Abbreviation",
                con);
            if (!CommectionOpen(ref con, ref error))
            {
                Debug.WriteLine("MYDEBUG: " + error);
                MessageBox.Show("Подключение к базе данных отсутсвует");
                return;

            }

            cmd.Parameters.AddWithValue("@Abbreviation", initial);

            reader = cmd.ExecuteReader();

            if (reader.HasRows)
            {
                if (reader.Read())
                {
                    chek = true;
                }
            }

            reader.Close();

            con.Close();
            if (chek == false)
            {
                cmd = new SQLiteCommand(
                    "Insert Into [Teacher] (FirstName, SurName, MiddleName, Abbreviation) Values(@FirstName, @SurName, @MiddleName, @Abbreviation)",
                    con);
                if (!CommectionOpen(ref con, ref error))
                {
                    Debug.WriteLine("MYDEBUG: " + error);
                    MessageBox.Show("Подключение к базе данных отсутсвует");
                    return;

                }

                cmd.Parameters.AddWithValue("@FirstName", firstName);
                cmd.Parameters.AddWithValue("@SurName", lastName);
                cmd.Parameters.AddWithValue("@MiddleName", middleName);
                cmd.Parameters.AddWithValue("@Abbreviation", initial);

                reader = cmd.ExecuteReader();


                reader.Close();

                con.Close();
                chek = false;
            }

            cmd = new SQLiteCommand(
                "SELECT [Teacher].[ID_Teacher] from [Teacher] where [Teacher].[Abbreviation] = @Abbreviation",
                con);
            if (!CommectionOpen(ref con, ref error))
            {
                Debug.WriteLine("MYDEBUG: " + error);
                MessageBox.Show("Подключение к базе данных отсутсвует");
                return;

            }

            cmd.Parameters.AddWithValue("@Abbreviation", initial);

            reader = cmd.ExecuteReader();

            if (reader.HasRows)
            {
                if (reader.Read())
                {
                    idTeacher = reader.GetInt32(0);
                }
            }

            reader.Close();

            con.Close();
            // проверка есть ли в бд, + заполнение. если в бд такой хуесос присуствует то проверяем заполнен ли он по группам если нет то заполняем и с преметами та же хуйня
            foreach (string value in subjectList)
            {
                cmd = new SQLiteCommand(
                    "select [Subject].[ID_Subject] from [Subject] where [Subject].[SubjectFullName] = @FullName",
                    con);

                if (!CommectionOpen(ref con, ref error))
                {
                    Debug.WriteLine("MYDEBUG: " + error);
                    MessageBox.Show("Подключение к базе данных отсутсвует");
                    return;

                }

                cmd.Parameters.AddWithValue("@FullName", value);

                reader = cmd.ExecuteReader();

                if (reader.HasRows)
                {
                    if (reader.Read())
                    {
                        idSubject = reader.GetInt32(0);
                    }
                }

                reader.Close();

                con.Close();

                cmd = new SQLiteCommand(
                    "SELECT ID_Sheet from Sheet where Teacher_ID = @ID_Teacher and Subject_ID = @ID_Subject", con);

                if (!CommectionOpen(ref con, ref error))
                {
                    Debug.WriteLine("MYDEBUG: " + error);
                    MessageBox.Show("Подключение к базе данных отсутсвует");
                    return;

                }

                cmd.Parameters.AddWithValue("@ID_Teacher", idTeacher);
                cmd.Parameters.AddWithValue("@ID_Subject", idSubject);

                reader = cmd.ExecuteReader();

                if (reader.HasRows)
                {
                    if (reader.Read())
                    {
                        tempSheet = reader.GetInt32(0);
                    }
                }

                reader.Close();

                con.Close();
                if (tempSheet == 0)
                {
                    cmd = new SQLiteCommand(
                        "Insert Into [Sheet] (Teacher_ID, Subject_ID) Values(@ID_Teacher, @ID_Subject) ", con);

                    if (!CommectionOpen(ref con, ref error))
                    {
                        Debug.WriteLine("MYDEBUG: " + error);
                        MessageBox.Show("Подключение к базе данных отсутсвует");
                        return;

                    }

                    cmd.Parameters.AddWithValue("@ID_Teacher", idTeacher);
                    cmd.Parameters.AddWithValue("@ID_Subject", idSubject);

                    reader = cmd.ExecuteReader();

                    reader.Close();

                    con.Close();
                }

                tempSheet = 0;
            }


            foreach (string value in groupList)
            {
                cmd = new SQLiteCommand(
                    "select [Groupes].[ID_Group] from [Groupes] where [Groupes].[GroupName] = @Groupes",
                    con);

                if (!CommectionOpen(ref con, ref error))
                {
                    Debug.WriteLine("MYDEBUG: " + error);
                    MessageBox.Show("Подключение к базе данных отсутсвует");
                    return;

                }

                cmd.Parameters.AddWithValue("@Groupes", value);

                reader = cmd.ExecuteReader();

                if (reader.HasRows)
                {
                    if (reader.Read())
                    {
                        idGroup = reader.GetInt32(0);
                    }
                }

                reader.Close();

                con.Close();

                cmd = new SQLiteCommand(
                    "SELECT ID_SheetGroup from SheetGroup where Teacher_ID = @ID_Teacher and Group_ID = @ID_Group", con);

                if (!CommectionOpen(ref con, ref error))
                {
                    Debug.WriteLine("MYDEBUG: " + error);
                    MessageBox.Show("Подключение к базе данных отсутсвует");
                    return;

                }

                cmd.Parameters.AddWithValue("@ID_Teacher", idTeacher);
                cmd.Parameters.AddWithValue("@ID_Group", idGroup);

                reader = cmd.ExecuteReader();

                if (reader.HasRows)
                {
                    if (reader.Read())
                    {
                        tempGroup = reader.GetInt32(0);
                    }
                }

                reader.Close();

                con.Close();
                if (tempGroup == 0)
                {
                    cmd = new SQLiteCommand(
                        "Insert Into [SheetGroup] (Teacher_ID, Group_ID) Values(@ID_Teacher, @ID_Group) ", con);

                    if (!CommectionOpen(ref con, ref error))
                    {
                        Debug.WriteLine("MYDEBUG: " + error);
                        MessageBox.Show("Подключение к базе данных отсутсвует");
                        return;

                    }

                    cmd.Parameters.AddWithValue("@ID_Teacher", idTeacher);
                    cmd.Parameters.AddWithValue("@ID_Group", idGroup);

                    reader = cmd.ExecuteReader();

                    reader.Close();

                    con.Close();
                }

                tempSheet = 0;
            }

        }





        public string SendInfo(Data writestring)
        {
            string sendInicial = writestring.inicial;
            return sendInicial;
        }

        public void WriteRange(Data writestring1)
        {
            object[,] writeString = writestring1.data;
            object[,] writeStringGroup = writestring1.dataGroup;
            object[,] writeString1Semestr = writestring1.data1Semestr;
            object[,] writeString2Semestr = writestring1.data2Semestr;
            object fioo = writestring1.fio;
            int starti = 9;
            int starty = 2;
            int endi = writestring1.end;
            int endy = 2;
            int startyGroup = 3;
            int endyGroup = 3;
            int starty1Semestr = 4;
            int endy1Semestr = 4;
            int starty2Semestr = 5;
            int endy2Semestr = 5;

            ExcelRange range = ws.Cells[starti, starty, endi, endy];
            range.Value = writeString;
            ExcelRange rangeGroup = ws.Cells[starti, startyGroup, endi, endyGroup];
            rangeGroup.Value = writeStringGroup;
            ExcelRange range1Semestr = ws.Cells[starti, starty1Semestr, endi, endy1Semestr];
            range1Semestr.Value = writeString1Semestr;
            ExcelRange range2Semestr = ws.Cells[starti, starty2Semestr, endi, endy2Semestr];
            range2Semestr.Value = writeString2Semestr;
            ExcelRange fiooo = ws.Cells[3, 2];
            fiooo.Value = fioo;
            Debug.WriteLine(ws.Cells[2, 4].Value); Debug.WriteLine(ws.Cells[4, 2].Value);


        }

        public void SaveAs(string path)
        {
            //wb.SaveAs(path);
            excel.SaveAs(path);
        }
        public void Close()
        {
            //wb.Close();
            fs.Close();
        }

    }
}


