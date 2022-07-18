using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows;

namespace FullPackMPTWrite.ScriptsWriteDiscription
{
    class ExcelWrite
    {
        public int g, v;
        public int numendi;
        string path = "";
        ExcelPackage excel;
        ExcelWorkbook wb;
        private ExcelWorksheet ws2;
        ExcelWorksheet ws;
        FileStream fs;

        public ExcelWrite(string path, string Sheet)
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                //выбор дока и листа
                this.path = path;
                fs = new FileStream(path, FileMode.Open);
                excel = new ExcelPackage(fs);
                wb = excel.Workbook;
                ws = wb.Worksheets[Sheet];
                ws2 = wb.Worksheets["Дисциплины"];
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                MessageBox.Show("Нет Нужного файла; " + e);
                throw;
            }
            //Лицензия

        }

        public List<int> FindWeekColumn(int _endColumn, string _week, int _colStart)
        {

            int colEnd = _endColumn;
            int rowStart = 8;
            List<int> cords = new List<int>();

            foreach (ExcelRangeBase value in ws.Cells[rowStart, _colStart, rowStart, colEnd])
            {
                if (value.Value.ToString() == _week)
                {
                    cords.Add(value.End.Column);
                }
            }

            return cords;
        }
        public List<int> FindGroupCells(string _group)
        {

            int colum = 3;
            int firstRowGroup = 9;
            ExcelRange cellsGroupTemp = ws.Cells[firstRowGroup, colum, 58, colum];

            object cellsGroup = cellsGroupTemp.Value;

            List<int> cordsGroup = new List<int>();

            foreach (ExcelRangeBase valueGroup in ws2.Cells[firstRowGroup, colum, 58, colum])
            {
                if (valueGroup.Value != null)
                {
                    if (valueGroup.Value.ToString() == _group)
                    {
                        cordsGroup.Add(valueGroup.End.Row);
                    }
                }
            }

            return cordsGroup;
        }

        public void WriteFirst(string _week, string _subject, string _group, int _endColumn, string _dataStart, bool _writeng)
        {
            List<int> weekColumn = FindWeekColumn(_endColumn, _week, 4);
            List<int> groupCells = FindGroupCells(_group);
            int columnSubject = 2;
            DateTime date = new DateTime();
            DateTime dateStart = new DateTime();
            bool isWrite = false;

            for (int i = 0; i <= groupCells.Count - 1; i++)
            {
                if (isWrite)
                    return;
                foreach (ExcelRangeBase valueCells in ws2.Cells[9, columnSubject, groupCells[i], columnSubject])
                {
                    if (isWrite)
                        return;

                    object temptest = valueCells.Value;
                    //Отделяем от ллишних слов (МДК, ПП, УП)
                    if (temptest is null)
                    {
                        continue;
                    }

                    var cell = new Cell(temptest, columnSubject, groupCells[i]);



                    string subject = _subject.Trim();
                    if (cell.Value == subject)
                    {
                        //gorupCell[i]
                        int mount;
                        string _mounth = ws.Cells["Q3"].Value.ToString();
                        _mounth = _mounth.Trim();
                        string year = ws.Cells[3, 21].Value.ToString();
                        switch (_mounth)
                        {
                            case "Сентябрь":
                                mount = 09;
                                break;
                            case "Октябрь":
                                mount = 10;
                                break;
                            case "Ноябрь":
                                mount = 11;
                                break;
                            case "Декабрь":
                                mount = 12;
                                break;
                            case "Январь":
                                mount = 01;
                                break;
                            case "Февраль":
                                mount = 02;
                                break;
                            case "Март":
                                mount = 03;
                                break;
                            case "Апрель":
                                mount = 04;
                                break;
                            case "Май":
                                mount = 05;
                                break;
                            case "Июнь":
                                mount = 06;
                                break;
                            case "Июль":
                                mount = 07;
                                break;
                            default:
                                mount = 0;
                                break;
                        }
                        string[] splitData = _dataStart.Split('.');
                        string dayStart = splitData[0];
                        string mounthStart = splitData[1];
                        string yearStart = splitData[2];
                        dateStart = new DateTime(Convert.ToInt32(yearStart), Convert.ToInt32(mounthStart), Convert.ToInt32(dayStart));
                        int dayHave = Convert.ToInt32(dayStart);
                        int startColumn = 0;

                        bool cheking = true;

                        while (cheking)
                        {
                            try
                            {
                                date = new DateTime(Convert.ToInt32(year), Convert.ToInt32(mount), Convert.ToInt32(dayHave));
                                cheking = false;
                            }
                            catch
                            {
                                dayHave--;
                                cheking = true;
                            }

                        }

                        if (date < dateStart)
                            return;
                        else
                        {
                            if (mount == Convert.ToInt32(mounthStart))
                            {
                                foreach (var dayValue in ws.Cells[7, 4, 7, _endColumn])
                                {
                                    if (Convert.ToInt32(dayValue.Value.ToString()) == Convert.ToInt32(dayStart))
                                        startColumn = dayValue.End.Row;
                                }

                                int rowClear = groupCells[i];
                                if (!_writeng)
                                {
                                    foreach (ExcelRangeBase cellValue in ws.Cells[rowClear, startColumn, rowClear, _endColumn])
                                    {

                                        if (cellValue.Style.Fill.BackgroundColor.Rgb != "FF008000")
                                        {
                                            cellValue.Value = "";
                                        }
                                    }
                                }
                                List<int> weekColumn1 = FindWeekColumn(_endColumn, _week, startColumn);
                                WriteCell(weekColumn1, groupCells, i);
                                isWrite = true;

                                cheking = true;
                            }
                            else
                            {
                                int rowClear = groupCells[i];
                                if (!_writeng)
                                {
                                    foreach (ExcelRangeBase cellValue in ws.Cells[rowClear, 4, rowClear, _endColumn])
                                    {

                                        if (cellValue.Style.Fill.BackgroundColor.Rgb != "FF008000")
                                        {
                                            cellValue.Value = "";
                                        }
                                    }
                                }
                                cheking = true;
                                WriteCell(weekColumn, groupCells, i);
                                isWrite = true;
                            }

                        }



                    }
                    else
                    {
                        Debug.WriteLine("Что то пошло не так");
                        //Debug.WriteLine("_group = " + _group, " groupCells = " + groupCells);
                    }

                }
            }

        }


        public void WriteIenum(string _week, string _subject, string _group, int _endColumn, bool _ienum, string _dataStart, bool _writeng)
        {
            List<int> weekColumn = FindWeekColumn(_endColumn, _week, 4);
            List<int> groupCells = FindGroupCells(_group);
            DateTime date = new DateTime();
            DateTime dateStart = new DateTime();
            bool isWrite = false;
            int columnSubject = 2;

            for (int i = 0; i <= groupCells.Count - 1; i++)
            {
                if (isWrite)
                    return;
                foreach (ExcelRangeBase valueCells in ws2.Cells[9, columnSubject, groupCells[i], columnSubject])
                {
                    if (isWrite)
                        return;
                    object temptest = valueCells.Value;
                    //Отделяем от ллишних слов (МДК, ПП, УП)
                    if (temptest is null)
                    {
                        continue;
                    }

                    var cell = new Cell(temptest, columnSubject, groupCells[i]);


                    string subject = _subject.Trim();
                    if (cell.Value == subject)
                    {
                        int mount;
                        string _mounth = ws.Cells["Q3"].Value.ToString();
                        _mounth = _mounth.Trim();
                        string year = ws.Cells[3, 21].Value.ToString();
                        switch (_mounth)
                        {
                            case "Сентябрь":
                                mount = 09;
                                break;
                            case "Октябрь":
                                mount = 10;
                                break;
                            case "Ноябрь":
                                mount = 11;
                                break;
                            case "Декабрь":
                                mount = 12;
                                break;
                            case "Январь":
                                mount = 01;
                                break;
                            case "Февраль":
                                mount = 02;
                                break;
                            case "Март":
                                mount = 03;
                                break;
                            case "Апрель":
                                mount = 04;
                                break;
                            case "Май":
                                mount = 05;
                                break;
                            case "Июнь":
                                mount = 06;
                                break;
                            case "Июль":
                                mount = 07;
                                break;
                            default:
                                mount = 0;
                                break;
                        }
                        string[] splitData = _dataStart.Split('.');
                        string dayStart = splitData[0];
                        string mounthStart = splitData[1];
                        string yearStart = splitData[2];
                        dateStart = new DateTime(Convert.ToInt32(yearStart), Convert.ToInt32(mounthStart), Convert.ToInt32(dayStart));
                        int dayHave = Convert.ToInt32(dayStart);
                        int startColumn = 0;

                        bool cheking = true;

                        while (cheking)
                        {
                            try
                            {
                                date = new DateTime(Convert.ToInt32(year), Convert.ToInt32(mount), Convert.ToInt32(dayHave));
                                cheking = false;
                            }
                            catch
                            {
                                dayHave--;
                                cheking = true;
                            }

                        }

                        if (date < dateStart)
                            return;
                        else
                        {
                            if (mount == Convert.ToInt32(mounthStart))
                            {
                                foreach (var dayValue in ws.Cells[7, 4, 7, _endColumn])
                                {
                                    if (Convert.ToInt32(dayValue.Value.ToString()) == Convert.ToInt32(dayStart))
                                        startColumn = dayValue.End.Row;
                                }

                                int rowClear = groupCells[i];
                                if (!_writeng)
                                {
                                    foreach (ExcelRangeBase cellValue in ws.Cells[rowClear, startColumn, rowClear, _endColumn])
                                    {

                                        if (cellValue.Style.Fill.BackgroundColor.Rgb != "FF008000")
                                        {
                                            cellValue.Value = "";
                                        }
                                    }
                                }
                                List<int> weekColumn1 = FindWeekColumn(_endColumn, _week, startColumn);
                                for (int j = 0; j <= weekColumn.Count - 1; j++)
                                {
                                    int x = groupCells[i];
                                    int y = weekColumn[j];

                                    bool ienum = CheckIenum(y);

                                    if (ienum && _ienum)
                                    {
                                        string hour = "";
                                        ExcelRange writeRange = ws.Cells[x, y];
                                        string color = ws.Cells[x, y].Style.Fill.BackgroundColor.Rgb;
                                        if (color != "FF008000")
                                        {
                                            if (writeRange.Value != null)
                                                hour = ws.Cells[x, y].Value.ToString();

                                            if (hour != "")
                                                writeRange.Value = Convert.ToInt32(hour) + 2;
                                            else
                                                writeRange.Value = 2;
                                        }
                                        else
                                        {
                                            Debug.WriteLine("Зеленая ячейка");
                                        }
                                        isWrite = true;

                                    }
                                    if (!ienum && !_ienum)
                                    {
                                        string hour = "";
                                        ExcelRange writeRange = ws.Cells[x, y];
                                        string color = ws.Cells[x, y].Style.Fill.BackgroundColor.Rgb;
                                        if (color != "FF008000")
                                        {
                                            if (writeRange.Value != null)
                                                hour = ws.Cells[x, y].Value.ToString();

                                            if (hour != "")
                                                writeRange.Value = Convert.ToInt32(hour) + 2;
                                            else
                                                writeRange.Value = 2;
                                        }
                                        else
                                        {
                                            Debug.WriteLine("Зеленая ячейка");
                                        }
                                        isWrite = true;
                                    }

                                }


                                cheking = true;
                            }
                            else
                            {
                                int rowClear = groupCells[i];
                                if (!_writeng)
                                {
                                    foreach (ExcelRangeBase cellValue in ws.Cells[rowClear, 4, rowClear, _endColumn])
                                    {

                                        if (cellValue.Style.Fill.BackgroundColor.Rgb != "FF008000")
                                        {
                                            cellValue.Value = "";
                                        }
                                    }
                                }
                                cheking = true;
                                for (int j = 0; j <= weekColumn.Count - 1; j++)
                                {
                                    int x = groupCells[i];
                                    int y = weekColumn[j];

                                    bool ienum = CheckIenum(y);

                                    if (ienum && _ienum)
                                    {
                                        string hour = "";
                                        ExcelRange writeRange = ws.Cells[x, y];
                                        string color = ws.Cells[x, y].Style.Fill.BackgroundColor.Rgb;
                                        if (color != "FF008000")
                                        {
                                            if (writeRange.Value != null)
                                                hour = ws.Cells[x, y].Value.ToString();

                                            if (hour != "")
                                                writeRange.Value = Convert.ToInt32(hour) + 2;
                                            else
                                                writeRange.Value = 2;
                                        }
                                        else
                                        {
                                            Debug.WriteLine("Зеленая ячейка");
                                        }
                                        isWrite = true;

                                    }
                                    if (!ienum && !_ienum)
                                    {
                                        string hour = "";
                                        ExcelRange writeRange = ws.Cells[x, y];
                                        string color = ws.Cells[x, y].Style.Fill.BackgroundColor.Rgb;
                                        if (color != "FF008000")
                                        {
                                            if (writeRange.Value != null)
                                                hour = ws.Cells[x, y].Value.ToString();

                                            if (hour != "")
                                                writeRange.Value = Convert.ToInt32(hour) + 2;
                                            else
                                                writeRange.Value = 2;
                                        }
                                        else
                                        {
                                            Debug.WriteLine("Зеленая ячейка");
                                        }

                                        //Debug.WriteLine(ws.Cells[x, y].Value);
                                        isWrite = true;
                                    }

                                }

                            }

                        }

                    }
                    else
                    {
                        Debug.WriteLine("Что то пошло не так");
                        // Debug.WriteLine("_group = " + _group, " groupCells = " + groupCells);
                    }


                    cell.Value = "";
                }

            }

        }
       

        private bool CheckIenum(int y)
        {
            bool ienum = false;
            string mont = ws.Cells[3, 17].Value.ToString().Trim();
            DateTime dateday = new DateTime(Convert.ToInt32(ws.Cells[3, 21].Value), GetMounth(mont), Convert.ToInt32(ws.Cells[7, y].Value));
            ienum = GetWeekType(dateday);
            return ienum;
        }

        private void WriteCell(List<int> weekColumn, List<int> groupCells, int i)
        {
            string hour = "";
            for (int j = 0; j <= weekColumn.Count - 1; j++)
            {
                int x = groupCells[i];
                int y = weekColumn[j];
                /*  ExcelRange write =*/


                ExcelRange writeRange = ws.Cells[x, y];
                string color = ws.Cells[x, y].Style.Fill.BackgroundColor.Rgb;
                if (color != "FF008000")
                {
                    if (writeRange.Value != null)
                        hour = ws.Cells[x, y].Value.ToString();

                    if (hour != "")
                        writeRange.Value = Convert.ToInt32(hour) + 2;
                    else
                        writeRange.Value = 2;
                }
                else
                {
                    Debug.WriteLine("Зеленая ячейка");
                }
                //Debug.WriteLine(ws.Cells[x, y].Value);
            }
        }
        public int GetMounth(string _month)
        {
            switch (_month)
            {
                case "Сентябрь":
                    return 09;
                case "Октябрь":
                    return 10;
                case "Ноябрь":
                    return 11;
                case "Декабрь":
                    return 12;
                case "Январь":
                    return 01;
                case "Февраль":
                    return 02;
                case "Март":
                    return 03;
                case "Апрель":
                    return 04;
                case "Май":
                    return 05;
                case "Июнь":
                    return 06;
                case "Июль":
                    return 07;
                default:
                    return 0;
            }
        }
        public void SaveAs(string path)
        {
            excel.SaveAs(path);
        }

        public void Close()
        {
            fs.Close();
        }

        public bool GetWeekType(DateTime _date)
        {
            int year = _date.Year;
            DateTime newDateSep = new DateTime(year, 09, 01);
            DateTime newDateJan = new DateTime(year, 01, 12);


            var calendar = CultureInfo.CurrentCulture.Calendar;
            var weekNumSep = calendar.GetWeekOfYear(newDateSep, CalendarWeekRule.FirstFullWeek, DayOfWeek.Monday);
            var weekNumJan = calendar.GetWeekOfYear(newDateJan, CalendarWeekRule.FirstFullWeek, DayOfWeek.Monday);
            var weekNum = calendar.GetWeekOfYear(_date, CalendarWeekRule.FirstFullWeek, DayOfWeek.Monday);

            var dayAmount = _date.DayOfYear;

            // Debug.WriteLine("Неделя сентября " + weekNumSep + "; Неделя новая " + weekNum);
            // Debug.WriteLine(dayAmount);
            if (dayAmount > 244)
            {
                if ((weekNum - weekNumSep) % 2 == 0)
                {
                    return true;
                }
            }
            else
            {
                if ((weekNum - weekNumJan) % 2 == 0)
                {
                    return true;
                }
            }
            return false;
        }

    }
    public class Cell
    {
        public AdressCell AdressCell { get; set; }
        public string Value { get; set; }

        public Cell(object _value, int _weight, int _height)
        {
            string finalResult = "";
            string tempValue = _value.ToString();
            string[] tempSplit = tempValue.Split(new[] { ' ' });
            if (tempSplit[0] == "МДК" || tempSplit[0] == "УП" || tempSplit[0] == "ПП")
            {
                var tempResult = tempSplit.Skip(2);

                foreach (var valueResult in tempResult)
                {
                    if (valueResult != null)
                        finalResult = finalResult + " " + valueResult;
                }

                tempValue = finalResult.Trim();
            }
            Value = tempValue;

            AdressCell = new AdressCell { Heigt = _height, Weight = _weight };

        }

    }

    public class AdressCell
    {
        public int Weight { get; set; }
        public int Heigt { get; set; }
    }
}
