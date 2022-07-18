using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;

using System.Data.SQLite;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Documents;

namespace FullPackMPTWrite.ScriptsWriteDiscription
{
    class Excel
    {



        SQLiteConnection con = new SQLiteConnection("Data Source=./MPTList.db");
        SQLiteDataAdapter da, da1;
        SQLiteCommand cmd;
        DataSet ds;
        DataTable dt;


        public int g, v;
        public int numendi;
        string path = "";
        ExcelPackage excel;
        ExcelWorkbook wb;
        ExcelWorksheet ws;
        FileStream fs;


        int groupCellRow = 9;
        string teacher = "";
        string subject = "";
        string tempGroup = "";
        private string group;

        public Excel(string path, int Sheet)
        {
            //Лицензия
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            //выбор дока и листа
            this.path = path;
            fs = new FileStream(path, FileMode.Open);
            excel = new ExcelPackage(fs);
            wb = excel.Workbook;
            ws = wb.Worksheets[Sheet];
        }
        public async void Work(string _folder)
        {
            //GetFiles(_folder);
            string folder = _folder;
            for (int i = 1; i < 4; i++)
            {
                int column = i * 3;
                int row = 8;
                string dataStart = System.Convert.ToString(ws.Cells[row, column].Value);
                string[] ddataStart = dataStart.Split(' ');
                dataStart = ddataStart[0];
                bool writeng = false;
                
                string week = "";
                for (int a = 0; a < 6; a++)
                {
                    switch (a)
                    {
                        case 0:
                            week = "Пн";
                            break;
                        case 1:
                            week = "Вт";
                            break;
                        case 2:
                            week = "Ср";
                            break;
                        case 3:
                            week = "Чт";
                            break;
                        case 4:
                            week = "Пт";
                            break;
                        case 5:
                            week = "Сб";
                            break;
                    }
                    Chekday(12 + 14 * a, 13 + 14 * a, column, folder, week, dataStart, writeng);
                    writeng = true;
                    Chekday(14 + 14 * a, 15 + 14 * a, column, folder, week, dataStart, writeng);
                    Chekday(16 + 14 * a, 17 + 14 * a, column, folder, week, dataStart, writeng);
                    Chekday(18 + 14 * a, 19 + 14 * a, column, folder, week, dataStart, writeng);
                    Chekday(20 + 14 * a, 21 + 14 * a, column, folder, week, dataStart, writeng);
                    Chekday(22 + 14 * a, 23 + 14 * a, column, folder, week, dataStart, writeng);
                    writeng = false;
                }

            }
            MessageBox.Show("Заполнение закончено");
        }


        public void Chekday(int _firstCell, int _secondCell, int _column, string _folder, string _week, string _dataStart, bool _writeng)
        {
            string week = _week;

            if (ws.Cells[_firstCell, _column].Value != null)
            {


                subject = ws.Cells[_firstCell, _column].Value.ToString();

                //ExcelRange cellSecond = ws.Cells[secondCell, column];
                teacher = (string)ws.Cells[_secondCell, _column].Value;
                // ExcelRange cellGroup = ws.Cells[groupCellRow, column];
                tempGroup = (string)ws.Cells[groupCellRow, _column].Value;
                object groupObject = ws.Cells[groupCellRow, _column].Value; 
                group = groupObject.ToString();


                string first = CheckFirtst(subject, teacher, _column);
                if (subject != "ПРАКТИКА")
                {
                    if (first != "yes")
                    {
                        string[] second = CheckSecond(subject, teacher, _column);
                        if (second[0] == "no")
                        {
                            string[] third = CheckThird(subject, teacher, _column);
                            bool ienum = false;
                            for (int f = 0; f <= 1; f++)
                            {
                                if (f == 0)
                                {
                                    teacher = third[1];
                                    subject = third[0];
                                    ienum = true;

                                }
                                else
                                {
                                    teacher = third[3];
                                    subject = third[2];
                                    ienum = false;
                                }

                                if (teacher != null)
                                {
                                    if (teacher != "")
                                    {
                                        if (subject != null)
                                        {
                                            if (subject != "")
                                            {
                                                string pathFile = _folder + @"\Ведомость учета часов " + teacher +
                                                                  ".xlsx";

                                                for (int i = 1; i < 12; i++)
                                                {
                                                    string mounth = "";
                                                    int endColumn = 0;

                                                    switch (i)
                                                    {
                                                        case 1:
                                                            mounth = "Сентябрь";
                                                            endColumn = 33;
                                                            break;
                                                        case 2:
                                                            mounth = "Октябрь";
                                                            endColumn = 34;
                                                            break;
                                                        case 3:
                                                            mounth = "Ноябрь";
                                                            endColumn = 33;
                                                            break;
                                                        case 4:
                                                            mounth = "Декабрь";
                                                            endColumn = 34;
                                                            break;
                                                        case 5:
                                                            mounth = "Январь";
                                                            endColumn = 34;
                                                            break;
                                                        case 6:
                                                            mounth = "Февраль";
                                                            endColumn = 31;
                                                            break;
                                                        case 7:
                                                            mounth = "Март";
                                                            endColumn = 34;
                                                            break;
                                                        case 8:
                                                            mounth = "Апрель";
                                                            endColumn = 33;
                                                            break;
                                                        case 9:
                                                            mounth = "Май";
                                                            endColumn = 34;
                                                            break;
                                                        case 10:
                                                            mounth = "Июнь";
                                                            endColumn = 33;
                                                            break;
                                                        case 11:
                                                            mounth = "Июль";
                                                            endColumn = 9;
                                                            break;

                                                    }


                                                    ExcelWrite ex = new ExcelWrite(pathFile, mounth);

                                                    ex.WriteIenum(week, subject, group, endColumn, ienum, _dataStart, _writeng);
                                                    ex.Close();
                                                    ex.SaveAs(pathFile);

                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            for (int nnn = 0; nnn <= 1; nnn++)
                            {
                                teacher = second[nnn];
                                string pathFile = _folder + @"\Ведомость учета часов " + teacher + ".xlsx";
                                for (int i = 1; i < 12; i++)
                                {
                                    string mounth = "";
                                    int endColumn = 0;

                                    switch (i)
                                    {
                                        case 1:
                                            mounth = "Сентябрь";
                                            endColumn = 33;
                                            break;
                                        case 2:
                                            mounth = "Октябрь";
                                            endColumn = 34;
                                            break;
                                        case 3:
                                            mounth = "Ноябрь";
                                            endColumn = 33;
                                            break;
                                        case 4:
                                            mounth = "Декабрь";
                                            endColumn = 34;
                                            break;
                                        case 5:
                                            mounth = "Январь";
                                            endColumn = 34;
                                            break;
                                        case 6:
                                            mounth = "Февраль";
                                            endColumn = 31;
                                            break;
                                        case 7:
                                            mounth = "Март";
                                            endColumn = 34;
                                            break;
                                        case 8:
                                            mounth = "Апрель";
                                            endColumn = 33;
                                            break;
                                        case 9:
                                            mounth = "Май";
                                            endColumn = 34;
                                            break;
                                        case 10:
                                            mounth = "Июнь";
                                            endColumn = 33;
                                            break;
                                        case 11:
                                            mounth = "Июль";
                                            endColumn = 9;
                                            break;

                                    }

                                    if (nnn == 0)
                                    {
                                        ExcelWrite exx = new ExcelWrite(pathFile, mounth);

                                        exx.WriteFirst(week, "Иностранный язык в профессиональной деятельности", group,
                                            endColumn, _dataStart, _writeng);
                                        exx.Close();
                                        exx.SaveAs(pathFile);
                                    }
                                    else
                                    {
                                        ExcelWrite exxx = new ExcelWrite(pathFile, mounth);

                                        exxx.WriteFirst(week, "Иностранный язык в профессиональной деятельности", group,
                                            endColumn, _dataStart, _writeng);
                                        exxx.Close();
                                        exxx.SaveAs(pathFile);
                                    }


                                }
                            }
                        }
                    }
                    else
                    {
                        string pathFile = _folder + @"\Ведомость учета часов " + teacher + ".xlsx";
                        for (int i = 1; i < 12; i++)
                        {
                            string mounth = "";
                            int endColumn = 0;

                            switch (i)
                            {
                                case 1:
                                    mounth = "Сентябрь";
                                    endColumn = 33;
                                    break;
                                case 2:
                                    mounth = "Октябрь";
                                    endColumn = 34;
                                    break;
                                case 3:
                                    mounth = "Ноябрь";
                                    endColumn = 33;
                                    break;
                                case 4:
                                    mounth = "Декабрь";
                                    endColumn = 34;
                                    break;
                                case 5:
                                    mounth = "Январь";
                                    endColumn = 34;
                                    break;
                                case 6:
                                    mounth = "Февраль";
                                    endColumn = 31;
                                    break;
                                case 7:
                                    mounth = "Март";
                                    endColumn = 34;
                                    break;
                                case 8:
                                    mounth = "Апрель";
                                    endColumn = 33;
                                    break;
                                case 9:
                                    mounth = "Май";
                                    endColumn = 34;
                                    break;
                                case 10:
                                    mounth = "Июнь";
                                    endColumn = 33;
                                    break;
                                case 11:
                                    mounth = "Июль";
                                    endColumn = 9;
                                    break;

                            }


                            ExcelWrite ex = new ExcelWrite(pathFile, mounth);

                            ex.WriteFirst(week, subject, group, endColumn, _dataStart, _writeng);
                            ex.Close();
                            ex.SaveAs(pathFile);

                        }

                    }

                }

            }
            else
            {
                if (ws.Cells[_secondCell, _column].Value != null)
                {
                    subject = null;
                    teacher = ws.Cells[_secondCell, _column].Value.ToString();
                    string[] third = CheckThird(subject, teacher, _column);
                    bool ienum = true;
                    for (int f = 0; f <= 1; f++)
                    {
                        if (f == 0)
                        {
                            teacher = third[1];
                            subject = third[0];
                            ienum = true;

                        }
                        else
                        {
                            teacher = third[3];
                            subject = third[2];
                            ienum = false;
                        }

                        if (teacher != null)
                        {
                            if (teacher != "")
                            {
                                if (subject != null)
                                {
                                    if (subject != "")
                                    {
                                        string pathFile = _folder + @"\Ведомость учета часов " + teacher +
                                                          ".xlsx";
                                        for (int i = 1; i < 12; i++)
                                        {
                                            string mounth = "";
                                            int endColumn = 0;

                                            switch (i)
                                            {
                                                case 1:
                                                    mounth = "Сентябрь";
                                                    endColumn = 33;
                                                    break;
                                                case 2:
                                                    mounth = "Октябрь";
                                                    endColumn = 34;
                                                    break;
                                                case 3:
                                                    mounth = "Ноябрь";
                                                    endColumn = 33;
                                                    break;
                                                case 4:
                                                    mounth = "Декабрь";
                                                    endColumn = 34;
                                                    break;
                                                case 5:
                                                    mounth = "Январь";
                                                    endColumn = 34;
                                                    break;
                                                case 6:
                                                    mounth = "Февраль";
                                                    endColumn = 31;
                                                    break;
                                                case 7:
                                                    mounth = "Март";
                                                    endColumn = 34;
                                                    break;
                                                case 8:
                                                    mounth = "Апрель";
                                                    endColumn = 33;
                                                    break;
                                                case 9:
                                                    mounth = "Май";
                                                    endColumn = 34;
                                                    break;
                                                case 10:
                                                    mounth = "Июнь";
                                                    endColumn = 33;
                                                    break;
                                                case 11:
                                                    mounth = "Июль";
                                                    endColumn = 9;
                                                    break;

                                            }


                                            ExcelWrite ex = new ExcelWrite(pathFile, mounth);

                                            ex.WriteIenum(week, subject, group, endColumn, ienum, _dataStart, _writeng);
                                            ex.Close();
                                            ex.SaveAs(pathFile);

                                        }
                                    }
                                }
                            }
                        }
                    }

                }
            }

        }


        //рабочая проверка препода и пары
        private string CheckFirtst(string subject, string teacher, int column)
        {
            string id_Teacher;
            string id_Subject;
            da = new SQLiteDataAdapter("SELECT ID_Teacher FROM [Teacher] WHERE Abbreviation  ='" + teacher + "'", con);
            da1 = new SQLiteDataAdapter("SELECT ID_Subject FROM Subject WHERE Subject ='" + subject + "'", con);
            dt = new DataTable();
            try
            {
                con.Open();
                da.Fill(dt);
                id_Teacher = dt.Rows[0].ToString();
                da1.Fill(dt);
                id_Subject = dt.Rows[0].ToString();
                con.Close();
                string returnstring = "yes";
                return returnstring;
            }
            catch
            {
                con.Close();
                // string[] returning = CheckSecond(subject, teacher, column);
                //return returning;
                string returnstring = "no";
                return returnstring;
            }

        }
        private string[] CheckSecond(string subject, string teacher, int column)
        {
            try
            {
                // надо починить
                if (teacher != null)
                {
                    string[] words = teacher.Split(new[] { ',' });
                    string firstTeacher = words[0];
                    string secondTeacher = words[1].Trim();


                    da = new SQLiteDataAdapter(
                        "SELECT ID_Teacher FROM Teacher WHERE Abbreviation  ='" + firstTeacher + "'", con);
                    da1 = new SQLiteDataAdapter(
                        "SELECT ID_Teacher FROM Teacher WHERE Abbreviation ='" + secondTeacher + "'", con);
                    dt = new DataTable();

                    con.Open();
                    da.Fill(dt);
                    string id_Teacher = dt.Rows[0].ToString();
                    da1.Fill(dt);
                    string id_TeacherSecond = dt.Rows[0].ToString();
                    con.Close();
                    string[] teachers = new string[2];
                    teachers[0] = firstTeacher;
                    teachers[1] = secondTeacher;
                    return teachers;


                }
                else
                {
                    string[] returnstring = new string[1];
                    returnstring[0] = "no";
                    return returnstring;
                }
            }
            catch (Exception)
            {
                string[] returnstring = new string[1];
                returnstring[0] = "no";
                return returnstring;
            }

        }
        private string[] CheckThird(string firstSubject, string secondSubject, int column)
        {
            string subjectNumerator = "";
            string teacherNumerator = "";
            string subjectDenominator = "";
            string teacherDenumerator = "";
            if (firstSubject != null)
            {
                string[] wordsNumerator = firstSubject.Split(new[] { ' ' });
                int amountWordsNumerator = wordsNumerator.Length;
                amountWordsNumerator = amountWordsNumerator - 3;
                teacherNumerator = wordsNumerator[amountWordsNumerator + 1] + " " +
                                          wordsNumerator[amountWordsNumerator + 2];

                for (int i = 0; i <= amountWordsNumerator; i++)
                {
                    subjectNumerator = subjectNumerator + " " + wordsNumerator[i];
                    subjectNumerator = subjectNumerator.Trim();
                }

                subjectNumerator.Trim();
            }
            //запись второго препода и пары
            if (secondSubject != null)
            {

                string[] wordsDenominator = secondSubject.Split(new[] { ' ' });
                int amountWordsDenominator = wordsDenominator.Length;
                amountWordsDenominator = amountWordsDenominator - 3;
                teacherDenumerator = wordsDenominator[amountWordsDenominator + 1] + " " +
                                            wordsDenominator[amountWordsDenominator + 2];

                for (int i = 0; i <= amountWordsDenominator; i++)
                {
                    subjectDenominator = subjectDenominator + " " + wordsDenominator[i];
                    subjectDenominator = subjectDenominator.Trim();
                }

                subjectDenominator.Trim();
            }

            string[] returnstring = new string[4];
            returnstring[0] = subjectNumerator;
            returnstring[1] = teacherNumerator;
            returnstring[2] = subjectDenominator;
            returnstring[3] = teacherDenumerator;
            return returnstring;
        }


    
        public void Close()
        {
            fs.Close();
        }
        string error = "";


        public void WorkPractic(string path)
        {
            int groupId = 0;

            SQLiteDataReader reader = null;

            List<int> rowEnd = GetEndRow();



            foreach (int value in rowEnd)
            {
                DateTime date1 = new DateTime();
                DateTime date2 = new DateTime();

               
                List<int> rowID = GetEndRow();
                List<int> teacherID = new List<int>();
                date1 = Convert.ToDateTime(ws.Cells[value, 4].Value.ToString());
                date2 = Convert.ToDateTime(ws.Cells[value, 5].Value.ToString());

                List<DateTime> dateArray = new List<DateTime>();
                dateArray = GetArrayDate(date1, date2);

                string groupName = ws.Cells[value, 3].Value.ToString();
                //string objectType = ws.Cells[value, 8].Value.ToString();

                groupId = GetGroupId(groupName, groupId);

                teacherID = TeacherId(groupId);

                for (int o = 0; o < teacherID.Count; o++)
                {
                    string abbreviation ="";
                    abbreviation = GetAbbreviation(teacherID, o);
                    string pathFile = path + @"\Ведомость учета часов " + abbreviation + ".xlsx";
                    for (int i = 1; i < 12; i++)
                    {
                        string mounth = "";
                        int endColumn = 0;
                        switch (i)
                        {
                            case 1:
                                mounth = "Сентябрь";
                                endColumn = 33;
                                break;
                            case 2:
                                mounth = "Октябрь";
                                endColumn = 34;
                                break;
                            case 3:
                                mounth = "Ноябрь";
                                endColumn = 33;
                                break;
                            case 4:
                                mounth = "Декабрь";
                                endColumn = 34;
                                break;
                            case 5:
                                mounth = "Январь";
                                endColumn = 34;
                                break;
                            case 6:
                                mounth = "Февраль";
                                endColumn = 31;
                                break;
                            case 7:
                                mounth = "Март";
                                endColumn = 34;
                                break;
                            case 8:
                                mounth = "Апрель";
                                endColumn = 33;
                                break;
                            case 9:
                                mounth = "Май";
                                endColumn = 34;
                                break;
                            case 10:
                                mounth = "Июнь";
                                endColumn = 33;
                                break;
                            case 11:
                                mounth = "Июль";
                                endColumn = 9;
                                break;

                        }

                        ExcelWritePractic exc = new ExcelWritePractic(pathFile, mounth);
                        exc.WritePract(groupName, dateArray, endColumn);
                        exc.Close();
                        exc.SaveAs(pathFile);
                        
                    }
                }
            }


        }

        private string GetAbbreviation(List<int> teacherID, int o)
        {
            SQLiteDataReader reader;
            string abbreviation = "";
            cmd = new SQLiteCommand(
                "select Abbreviation from [Teacher] where ID_Teacher = @TeacherID",
                con);

            if (!CommectionOpen(ref con, ref error))
            {
                Debug.WriteLine("MYDEBUG: " + error);
                MessageBox.Show("Подключение к базе данных отсутсвует");
                //return;
            }

            cmd.Parameters.AddWithValue("@TeacherID", teacherID[o]);

            reader = cmd.ExecuteReader();

            if (reader.HasRows)
            {
                if (reader.Read())
                {
                    int readerRowMax = reader.StepCount;
                    for (int i = 0; i < readerRowMax; i++)
                    {
                        abbreviation = reader.GetString(0);
                    }
                }
            }

            reader.Close();

            con.Close();
            return abbreviation;
        }

        private List<int> TeacherId(int groupId )
        {
            List<int> teacherID = new List<int>();
            SQLiteDataReader reader;
            cmd = new SQLiteCommand(
                "select Teacher_ID from [SheetGroup] where Group_ID = @GroupID",
                con);

            if (!CommectionOpen(ref con, ref error))
            {
                Debug.WriteLine("MYDEBUG: " + error);
                MessageBox.Show("Подключение к базе данных отсутсвует");
                //return;
            }

            cmd.Parameters.AddWithValue("@GroupID", groupId);

            reader = cmd.ExecuteReader();
            //
            {
                if (reader.Read())
                {
                    int readerRowMax = reader.StepCount;
                    for (int i = 0; i < readerRowMax; i++)
                    {
                        teacherID.Add(reader.GetInt32(i));
                    }
                }
            }

            reader.Close();
            string abbreviation = "";
            con.Close();
            return teacherID;
        }

        private int GetGroupId(string groupName, int groupId)
        {
            cmd = new SQLiteCommand(
                "select ID_Group from [Groupes] where GroupName = @GroupNamee",
                con);

            if (!CommectionOpen(ref con, ref error))
            {
                Debug.WriteLine("MYDEBUG: " + error);
                MessageBox.Show("Подключение к базе данных отсутсвует");
               // return;

            }
            SQLiteDataReader reader;
            cmd.Parameters.AddWithValue("@GroupNamee", groupName);

            reader = cmd.ExecuteReader();

            if (reader.HasRows)
            {
                if (reader.Read())
                {
                    groupId = reader.GetInt32(0);
                }
            }

            reader.Close();

            con.Close();
            return groupId;
        }

        public List<DateTime> GetArrayDate(DateTime date1, DateTime date2)
        {
            TimeSpan diff1 = date2.Subtract(date1);
            DateTime date = new DateTime();
            date = date1;
            List<DateTime> dateArray = new List<DateTime>();
            for (int i = 1; i <= Convert.ToInt32(diff1.Days + 1); i++)
            {
                dateArray.Add(date);
                Console.WriteLine(date.ToString().Substring(0, date.ToString().Length - 8));
                date = date.AddDays(1);
            }

            return dateArray;
        }

        public List<int> GetEndRow()
        {
            int row = 4;
            List<int> rowArray = new List<int>();
            while (ws.Cells[row, 1].Value != null)
            {
                rowArray.Add(row);
                row++;

            }
            return rowArray;
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