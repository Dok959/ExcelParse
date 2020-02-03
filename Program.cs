using System;
using System.Drawing;
using System.IO;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace App
{
    class Program
    {
        static void Main(string[] args)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWB;
            Excel.Worksheet xlSht;

            xlApp = new Excel.Application();
            var dir = Directory.GetCurrentDirectory();
            xlWB = xlApp.Workbooks.Open(dir + "\\Учет рабочего времени ноябрь 2019.xls");    
            xlSht = xlWB.ActiveSheet;

            int iLastRow = xlSht.Cells[xlSht.Rows.Count, "A"].End[Excel.XlDirection.xlUp].Row-1;
            int iLastCol = xlSht.UsedRange.Columns.Count;

            var arrData = (object[,])xlSht.Range["A1:Z" + iLastRow].Value;
            xlWB.Close(false);
            xlApp.Quit();

            FileInfo fileInf = new FileInfo(dir + "\\Parse.xlsx");
            if (fileInf.Exists)
            {
                fileInf.Delete();
            }

            var outbook = xlApp.Workbooks.Add();
            int rtime = GetWorkingDays(Convert.ToDateTime(arrData[2, 16].ToString().Substring(3,11).Replace(".", "-")),
                Convert.ToDateTime(arrData[2, 16].ToString().Substring(17, 10).Replace(".", "-")));

            outbook.ActiveSheet.Range["A9"].Value = "Табель учета рабочего времени с "
                + RTime(Convert.ToDateTime(arrData[2, 16].ToString().Substring(3, 11).Replace(".", "-")),
                arrData[2, 16].ToString().Substring(17, 2));
            Excel.Range _excelCells1 = (Excel.Range)outbook.ActiveSheet.Range["A9", "J9"].Cells;
            _excelCells1.Merge(Type.Missing);
            _excelCells1.Cells.Font.Name = "Times New Roman";
            _excelCells1.Cells.Font.Bold = 12;
            _excelCells1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            _excelCells1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            outbook.ActiveSheet.Range["A4"].Value = "Приложение №2";
            _excelCells1 = (Excel.Range)outbook.ActiveSheet.Range["A4", "J4"].Cells;
            _excelCells1.Merge(Type.Missing);
            _excelCells1.Cells.Font.Bold = 12;

            outbook.ActiveSheet.Range["A5"].Value = "К Приказу № 22 от 30.03.2018 г.";
            _excelCells1 = (Excel.Range)outbook.ActiveSheet.Range["A5", "J5"].Cells;
            _excelCells1.Merge(Type.Missing);
            _excelCells1.Cells.Font.Bold = 12;

            outbook.ActiveSheet.Range["A6"].Value = "«О изменении режима рабочего времени и ";
            _excelCells1 = (Excel.Range)outbook.ActiveSheet.Range["A6", "J6"].Cells;
            _excelCells1.Merge(Type.Missing);
            _excelCells1.Cells.Font.Size = 12;

            outbook.ActiveSheet.Range["A7"].Value = "применении взысканий за нарушение";
            _excelCells1 = (Excel.Range)outbook.ActiveSheet.Range["A7", "J7"].Cells;
            _excelCells1.Merge(Type.Missing);
            _excelCells1.Cells.Font.Size = 12;

            outbook.ActiveSheet.Range["A8"].Value = " трудовой дисциплины ООО «Прогресс»";
            _excelCells1 = (Excel.Range)outbook.ActiveSheet.Range["A8", "J8"].Cells;
            _excelCells1.Merge(Type.Missing);
            _excelCells1.Cells.Font.Size = 12;
            Excel.Range range3 = outbook.ActiveSheet.Range(outbook.ActiveSheet.Cells[4, 1], outbook.ActiveSheet.Cells[8, 10]);
            range3.Cells.Font.Name = "Times New Roman";
            range3.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            range3.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;

            int k = 0;
            int t = 13;
            int n = 0;
            int o = 0;
            bool flag = false;
            bool flag1 = false;
            outbook.ActiveSheet.Range["A11"].Value = "№ п/п";
            _excelCells1 = (Excel.Range)outbook.ActiveSheet.Range["A11", "A12"].Cells;
            _excelCells1.Merge(Type.Missing);
            outbook.ActiveSheet.Range["B11"].Value = "Подразделение";
            _excelCells1 = (Excel.Range)outbook.ActiveSheet.Range["B11", "B12"].Cells;
            _excelCells1.Merge(Type.Missing);
            outbook.ActiveSheet.Range["C11"].Value = "ФИО";
            _excelCells1 = (Excel.Range)outbook.ActiveSheet.Range["C11", "C12"].Cells;
            _excelCells1.Merge(Type.Missing);
            outbook.ActiveSheet.Range["D11"].Value = "Должность";
            _excelCells1 = (Excel.Range)outbook.ActiveSheet.Range["D11", "D12"].Cells;
            _excelCells1.Merge(Type.Missing);
            outbook.ActiveSheet.Range["E11"].Value = "Штатное \nколичество";
            outbook.ActiveSheet.Range["E12"].Value = "дней";
            outbook.ActiveSheet.Range["F11"].Value = "Фактическое количество по \nСКУД";
            _excelCells1 = (Excel.Range)outbook.ActiveSheet.Range["F11", "G11"].Cells;
            _excelCells1.Merge(Type.Missing);
            outbook.ActiveSheet.Range["F12"].Value = "дней";
            outbook.ActiveSheet.Range["G12"].Value = "часов";
            outbook.ActiveSheet.Range["H11"].Value = "Дельта между \nСКУД и штатными \nпоказателями";
            _excelCells1 = (Excel.Range)outbook.ActiveSheet.Range["H11", "H12"].Cells;
            _excelCells1.Merge(Type.Missing);
            outbook.ActiveSheet.Range["I11"].Value = "Примечание";
            _excelCells1 = (Excel.Range)outbook.ActiveSheet.Range["I11", "I12"].Cells;
            _excelCells1.Merge(Type.Missing);
            outbook.ActiveSheet.Range["J11"].Value = "Заключение \nруководителя по \nвыплате";
            _excelCells1 = (Excel.Range)outbook.ActiveSheet.Range["J11", "J12"].Cells;
            _excelCells1.Merge(Type.Missing);
            Excel.Range range2 = outbook.ActiveSheet.Range(outbook.ActiveSheet.Cells[11, 1], outbook.ActiveSheet.Cells[12, 10]);
            range2.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous;
            range2.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous;
            range2.Borders.get_Item(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlContinuous;
            range2.Borders.get_Item(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlContinuous;
            range2.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous;
            range2.Borders.Weight = 3;
            range2.Cells.Font.Name = "Times New Roman";
            range2.Cells.Font.Size = 12;
            range2.Cells.Font.Bold = 12;
            range2.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(0xD9, 0xD9, 0xD9));


            for (int iRow = 1; iRow <= iLastRow; iRow++)
            {
                flag1 = false;
                for (int iCol = 1; iCol <= iLastCol; iCol++) {
                    if (arrData[iRow, iCol] != null && arrData[iRow, iCol].ToString().Contains("Сотрудник"))
                    {
                        k += 1;
                        outbook.ActiveSheet.Range[$"A{t}"].Value = k;
                        var name = arrData[iRow, 11].ToString().Split();
                        string res;
                        if (name.Length == 3) { 
                            res = name[0] + " " + name[1].FirstOrDefault() + ". " + name[2].FirstOrDefault() + ".";}
                        else
                        {
                            res = name[0] + " " + name[1].FirstOrDefault() + ". ";
                        }
                        outbook.ActiveSheet.Range[$"C{t}"].Value = res;
                        outbook.ActiveSheet.Range[$"E{t}"].Value = rtime;
                    }
                    else if (arrData[iRow, iCol] != null && arrData[iRow, iCol].ToString().Contains("Отдел"))
                    {
                        outbook.ActiveSheet.Range[$"B{t}"].Value = arrData[iRow, 11].ToString();
                    }
                    else if(arrData[iRow, iCol] != null && arrData[iRow, iCol].ToString().Contains("Должность"))
                    {
                        if (arrData[iRow, 11] != null)
                            outbook.ActiveSheet.Range[$"D{t}"].Value = arrData[iRow, 11].ToString();
                    }
                    if (arrData[iRow, iCol] != null && arrData[iRow, iCol].ToString().Contains("Приход"))
                    {
                        n = 0;
                        o = 0;
                        flag = true;
                        flag1 = true;
                    }
                    else if (arrData[iRow, iCol] != null && flag == true 
                        && flag1 == false && arrData[iRow, iCol].ToString().Contains("-") == false
                        && arrData[iRow, iCol].ToString().Contains("Итого") == false)
                    {
                        var a = arrData[iRow, iCol];
                        if (arrData[iRow, 4].ToString().Contains("-") != true)
                            n++;
                        flag1 = true;
                        if (arrData[iRow, 18].ToString().Contains("-") != true) {
                        outbook.ActiveSheet.Range[$"I{t}"].Value += arrData[iRow, 1].ToString()
                            + " (" + arrData[iRow, 18].ToString() + ") - опоздание\n";
                            o++;
                        }
                    }
                    if (arrData[iRow, iCol] != null && arrData[iRow, iCol].ToString().Contains("Итого"))
                    {
                        flag = false;
                        outbook.ActiveSheet.Range[$"F{t}"].Value = n;
                        outbook.ActiveSheet.Range[$"G{t}"].Value = double.Parse(arrData[iRow, 10].ToString().Replace(":",","));
                        string temp1 = outbook.ActiveSheet.Range[$"I{t}"].Value;
                        if (temp1 != null)
                            outbook.ActiveSheet.Range[$"I{t}"].Value = temp1.Trim();
                        else
                            outbook.ActiveSheet.Range[$"I{t}"].Value = "";
                        double time = rtime * 7.45;
                        TimeSpan time1 = TimeSpan.FromHours(time / 10) + TimeSpan.FromMinutes(time % 10);
                        var temp = arrData[iRow, 10].ToString().Replace(":"," ").Split();
                        TimeSpan time2 = TimeSpan.FromHours(int.Parse(temp[0]))
                            + TimeSpan.FromMinutes(int.Parse(temp[1]));
                        if (time1 > time2)
                            outbook.ActiveSheet.Range[$"H{t}"].Value = (time1-time2).ToString() + " \n(переработка)";
                        else
                            outbook.ActiveSheet.Range[$"H{t}"].Value = (time2 - time1).ToString() + " \n(недоработка)";
                        if (o > 0)
                            outbook.ActiveSheet.Range[$"J{t}"].Value = "Штраф за " + o +" опоздание(ий)";
                        else
                            outbook.ActiveSheet.Range[$"J{t}"].Value = "Без штрафа";
                        t++;
                    }
                }
            }

            Excel.Range range1 = outbook.ActiveSheet.Range(outbook.ActiveSheet.Cells[13, 1],
                outbook.ActiveSheet.Cells[outbook.ActiveSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row,
                outbook.ActiveSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column]);
            range1.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous;
            range1.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous;
            range1.Borders.get_Item(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlContinuous;
            range1.Borders.get_Item(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlContinuous;
            range1.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous;
            range1.Cells.Font.Name = "Times New Roman";
            range1.Cells.Font.Size = 12;
            range1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            range1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            range2.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            range2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            range2.EntireColumn.AutoFit();

            Excel.Range range4 = outbook.ActiveSheet.Range(outbook.ActiveSheet.Cells[13, 2],
                outbook.ActiveSheet.Cells[outbook.ActiveSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row,4]);
            range4.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            range4 = outbook.ActiveSheet.Range(outbook.ActiveSheet.Cells[13, 2],
                outbook.ActiveSheet.Cells[outbook.ActiveSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row, 2]);
            range4.Cells.Font.Italic = 12;
            range4 = outbook.ActiveSheet.Range(outbook.ActiveSheet.Cells[11, 4],
                outbook.ActiveSheet.Cells[outbook.ActiveSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row, 4]);
            range4.WrapText = true;
            outbook.ActiveSheet.Columns["D"].ColumnWidth = 20;
            outbook.ActiveSheet.Columns["E"].ColumnWidth = 15;
            outbook.ActiveSheet.Columns["F"].ColumnWidth = 20;
            outbook.ActiveSheet.Columns["G"].ColumnWidth = 20;
            range2.EntireRow.RowHeight = 35;


            outbook.SaveAs(dir + "\\Parse.xlsx");

            outbook.Close();
            xlApp.Quit();
            Console.WriteLine("Программа отработала, нажмите любую кнопку");
            Console.ReadLine();
        }

        public static int GetWorkingDays(DateTime from, DateTime to)
        {
            var dayDifference = (int)to.Subtract(from).TotalDays;
            return Enumerable
                .Range(1, dayDifference)
                .Select(x => from.AddDays(x))
                .Count(x => x.DayOfWeek != DayOfWeek.Saturday && x.DayOfWeek != DayOfWeek.Sunday);
        }

        public static string RTime(DateTime from, string to)
        {
            string[] months = { "января", "февраля", "марта", "апреля",
                "мая", "июня", "июля", "августа", "сентября",
                "октября", "ноября", "декабря" };
            int mon = int.Parse(from.ToString().Substring(3, 2));
            string month = "";
            for (int i = 0; i <= mon; i++)
            {
                if (i ==mon)
                    month = months[i-1];
            }

            string first = int.Parse(from.ToString().Substring(3, 2)).ToString();
            string last = int.Parse(to).ToString();
            return first + " по " + last + " " + month;
        }
    }
}
