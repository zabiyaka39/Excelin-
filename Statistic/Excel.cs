using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using OfficeOpenXml;
using System.Globalization;
using System.Linq;
using OfficeOpenXml.Style;
using System.Drawing;
using OfficeOpenXml.Drawing.Chart;

namespace Statistic
{
    class Excel
    {
        // дата из csv
        string date;
       
        // список с типом данных датастатистик
        public List<DataStatiscics> dataStatiscics;


        public Excel(List<DataStatiscics> dataStatiscics, string date )
        {
            this.dataStatiscics = dataStatiscics;
            this.date = date;
  
            UpdateCurrentFile(); 
          

        }

        // ф-ция заполнения excel файла
        public void UpdateCurrentFile()
        {
            // выбираю текущий регион создаю переменные с датами для последующих действий 
            CultureInfo ci = new CultureInfo("ru-RU");
            int yearCurrent = DateTime.Parse(date).Year;
            int dateCurrent = DateTime.Parse(date).Day;
            int monthCurrent = DateTime.Parse(date).Month;
            string currentMonth = ci.DateTimeFormat.GetMonthName(monthCurrent);

            // обязательный пункт для работы с пакетом(условие лицензии)
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // путь к excel файлу исходя из требований csv
            string pathExcelFile = Environment.CurrentDirectory + $"\\excelMonth\\{currentMonth}.xlsx";

            // Проверяем наличие файла в дерриктории currentMonth
            FileInfo excelFile = new FileInfo(pathExcelFile);
            if (excelFile.Exists)
            {
                // Открывается файл, создается функция экземпляр класса для работы с файлом эксцель excel
                using (ExcelPackage excel = new ExcelPackage(excelFile))
                {
                    // Открывается Лист, создается функция экземпляр класса для работы с листом
                    ExcelWorksheet excelSheet = excel.Workbook.Worksheets.First();
                    excelSheet.Cells[excelSheet.Dimension.Address].AutoFitColumns();

                    // проверяем актуальност заголовка excel файла, сопоставляем ячейки в csv и ячейки в excel. при совпадении ячеек 
                    // находим актуальную дату и туда записываем
                    if ((string)(excelSheet.Cells[1, 1].Value) == currentMonth)
                    {
                        var end = excelSheet.Dimension.End;
                        int dataStatiscicsCount = dataStatiscics.Count;

                        foreach (DataStatiscics border in dataStatiscics)
                        {
                            for (int i = 2; i <= (end.Row); i++)
                            {
                                if ((string)excelSheet.Cells[i, 3].Value != null && (string)excelSheet.Cells[i, 3].Value == border.border)
                                {
                                    for (int column = 8; column <= end.Column; column++)
                                    {
                                        if (excelSheet.Cells[3, column].Value != null && Convert.ToInt32(excelSheet.Cells[3, column].Value) == dateCurrent)
                                        {
                                            excelSheet.Cells[i, column].Value = Convert.ToInt32(border.detections);
                                            break;
                                        }
                                       
                                    }
                                    break;
                                }
                            }
                        }

                        Console.WriteLine($"В таблицу за {currentMonth} введены данные за {date}");

                    }

                    excel.SaveAs(excelFile);
                    //сохраняем
                }

            }
            else
            {
                //Создается новый файл с текущим месяцем. Копируются все столбцы и строки из шаблона

                using (FileStream filestream = new FileStream(Environment.CurrentDirectory + $"\\excelMonth\\{currentMonth}.xlsx", FileMode.OpenOrCreate, FileAccess.Write, FileShare.ReadWrite))
                {
                    FileStream fs = new FileStream(Environment.CurrentDirectory + "\\Exceltemp.xlsx", FileMode.Open, FileAccess.Read);
                    byte[] array = new byte[fs.Length];
                    int bytesRead = -1;
                    while ((bytesRead = fs.Read(array, 0, array.Length)) > 0)
                    {
                        filestream.Write(array, 0, bytesRead);
                    }
                }
                // Открывается файл, создается функция экземпляр класса для работы с файлом эксцель excel 
                using (ExcelPackage excel = new ExcelPackage(excelFile))
                {
                    // Выбирается первый лист
                    ExcelWorksheet excelSheet = excel.Workbook.Worksheets.First();
                    // Объединяется в шапке несколько ячеек для записи текущего месяца 
                    excelSheet.Cells["A1:F1"].Merge = true;
                    excelSheet.Cells["A1:F1"].Value = currentMonth;
                    var start = excelSheet.Dimension.Start;
                    var end = excelSheet.Dimension.End;
                    var emptyEnd = end; 
                    
                    // Определят какой день недели первый день месяца  
                    DateTime dt = DateTime.Parse($"1/{monthCurrent}/{yearCurrent}");
                    int firstday = (int)dt.DayOfWeek - 1;
                    string[] weekDate = { "Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс" };
                    if (firstday != 0)
                    {
                        foreach (string dat in weekDate[0..firstday])
                        {
                            excelSheet.Cells[2, end.Column+1].Value = dat;
                            excelSheet.Column(end.Column+1).Hidden = true;
                            end = excelSheet.Dimension.End;
                        }
                    }
                  

                    // Сопоставляется количество дней в текущем месяце и записывается в таблицу вместе с днями недели 
                    int totalDays = DateTime.DaysInMonth(yearCurrent, monthCurrent);
                    for (int count = 1; count <= (totalDays); count++ )
                    {
                        DateTime dtime = DateTime.Parse($"{count}/{monthCurrent}/{yearCurrent}");

                        excelSheet.Cells[2,end.Column + count].Value = ci.DateTimeFormat.GetShortestDayName(dtime.DayOfWeek);
                        excelSheet.Cells[3,end.Column + count].Value = count;

                    }
                    end = excelSheet.Dimension.End;

                    // если последний день месяца воскресенье -ок, если нет, создает и скрывает пустые ячейки чтобы нормально построилась таблица 
                    if ((string)excelSheet.Cells[2, end.Column].Value != "Вс")
                    {
                        int ind = Array.IndexOf(weekDate, (string)excelSheet.Cells[2, end.Column].Value);

                        foreach (string dat in weekDate[0..ind])
                        {
                            excelSheet.Cells[2, end.Column + 1].Value = dat;
                            excelSheet.Column(end.Column+1).Hidden = true;
                            end = excelSheet.Dimension.End;
                        }

                    }
                    //Делаю границы таблицы
                    string modelRange = $"{start.Address}:{end.Address}";
                    using (ExcelRange range = excelSheet.Cells[modelRange]) 
                    {

                        range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Top.Color.SetColor(Color.Black);
                        range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Left.Color.SetColor(Color.Black);
                        range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Right.Color.SetColor(Color.Black);
                        range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Bottom.Color.SetColor(Color.Black);
                        
                    }

                    end = excelSheet.Dimension.End;
                    for (int index = 5; index <= end.Row; index++ )
                    {
                       
                       if (excelSheet.Cells[index,3].Value != null)
                       {
                            ExcelLineChart lineChart = excelSheet.Drawings.AddLineChart((string)excelSheet.Cells[index,3].Value, (eLineChartType)eChartType.LineMarkers);
                            DateTime dtime = DateTime.Parse($"1/{monthCurrent}/{yearCurrent}");
                            
                            int shift = 0;
                            int indexlinexhart = 0;

                            for (int d = 1; d <= end.Column-emptyEnd.Column; d++)
                            {
                                
                                if (d%7==0)
                                {
                                   
                                    var range1 = excelSheet.Cells[index, emptyEnd.Column+1 + shift, index, emptyEnd.Column+7 + shift];
                                    var range2 = excelSheet.Cells[2, emptyEnd.Column + 1, 2, emptyEnd.Column + 7];
                                    
                                    lineChart.Series.Add(range1, range2);
                            
                                    lineChart.Series[indexlinexhart].Header = $"Неделя {indexlinexhart+1}";
                                    shift += 7;
                                    indexlinexhart += 1;
                                }
                                
                            }
                            
                            lineChart.SetSize(500, 25);
                            lineChart.SetPosition(index-1 , 0, end.Column, 0);
                            lineChart.Title.Text = (string)excelSheet.Cells[index,3].Value;
                            lineChart.ShowHiddenData = true;
                        }
                        
                    }

                    // файл excel сохраняется и функция запускается рекурсивно  
                    excel.SaveAs(excelFile);
                    UpdateCurrentFile();

                }
            }
        }
     


    }
}
