using System;
using System.Collections.Generic;
using System.IO;

using System.Windows;
using System.Windows.Media;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Color = System.Drawing.Color;

namespace FullPackMPTWrite.ScriptsWriteDiscription
{
    class ExcelWritePractic
    {
        public int g, v;
        public int numendi;
        string path = "";
        ExcelPackage excel;
        ExcelWorkbook wb;
        private ExcelWorksheet ws2;
        ExcelWorksheet ws;
        FileStream fs;

        public ExcelWritePractic(string path, string Sheet)
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

        public void WritePract(string _groupName, List<DateTime> _dateArray, int _endColumn)
        {
            int mount;
            string _mounth = ws.Cells["Q3"].Value.ToString();
            _mounth = _mounth.Trim();
            mount = GetMounth(_mounth);

            string year = ws.Cells[3, 21].Value.ToString();

            List<int> groupRow = FindGroupCells(_groupName);
            DateTime date = new DateTime();

            foreach (var value in _dateArray)
            {
                string datatemp = value.ToString();
                string[] ddataStart = datatemp.Split(' ');
                datatemp = ddataStart[0];

                string[] splitData = datatemp.Split('.');
                string dayStart = splitData[0];
                string mounthStart = splitData[1];
                string yearStart = splitData[2];

                int dayHave = Convert.ToInt32(dayStart);

                //dateStart = new DateTime(Convert.ToInt32(yearStart), Convert.ToInt32(mount), Convert.ToInt32(dayStart));
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
                if (date < value)
                    return;
                else
                {
                    if (value.Month == mount)
                    {
                        bool chek = true;
                        chek = true;
                        foreach (var dayValue in ws.Cells[7, 4, 7, _endColumn])
                        {
                            if (chek)
                            {
                                if (value.Day == Convert.ToInt32(dayValue.Value.ToString()))
                                {
                                    foreach (int valueGroupRow in groupRow)
                                    {
                                        int column = dayValue.End.Column;
                                        ws.Cells[valueGroupRow, column].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        ws.Cells[valueGroupRow, column].Style.Fill.BackgroundColor.SetColor(Color.Green);
                                        ws.Cells[valueGroupRow, column].Value = "x";
                                    }

                                    chek = false;

                                }
                            }

                            //return;
                        }
                    }
                }

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
    }
}
