using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Time_Sheet_Constructor.Model
{
    /// <summary>
    /// Лист Чистовик
    /// </summary>
    public class ExportFair
    {
        const string tableLayoutPath =
            @"C:\Users\vadim.turetskiy\Documents\Табель\Time sheet constructor\Табель Шаблон.xlsx";

        private const string outputName =
            @"C:\Users\vadim.turetskiy\Documents\Табель\Time sheet constructor\Табель Выход.xlsx";

        const string fairSheetName = "Чистовик";

        /// <summary>
        /// Данные файла
        /// </summary>
        static FileInfo Fi => new FileInfo(tableLayoutPath);

        static ExcelPackage Excel => new ExcelPackage(Fi);

        private static int FirstRow = 12;
        private static int FirstColumn = 1;

        private static int FirstDay = 1;
        private static int LastDay = 15;

        private static int FirstDayColumn = 4;
        private static int LastDayColumn = 34;
       
        public static void Write(List<Person> persons)
        {
            using (var wb = Excel)
            {
                var row = FirstRow;
                var id = 1;

                foreach (var person in persons)
                {
                    var scheduleDay = 0;

                    if (person.EmployeeId == 0)
                    {
                        continue;
                    }
                    
                    wb.Workbook.Worksheets[fairSheetName].Cells[row, FirstColumn].Value = id;
                    wb.Workbook.Worksheets[fairSheetName].Cells[row, FirstColumn + 1].Value = person.GetFullName();
                    wb.Workbook.Worksheets[fairSheetName].Cells[row, FirstColumn + 2].Value = person.EmployeeId;

                    for (var column = FirstDayColumn; column <= LastDayColumn; column++)
                    {
                        if (column == 19)
                        {
                            continue;
                        }

                        if (person.Schedule[scheduleDay].AllWorkTime != 0)
                        {
                            wb.Workbook.Worksheets[fairSheetName].Cells[row + 3, column].Value += "Я";
                            wb.Workbook.Worksheets[fairSheetName].Cells[row + 1, column].Value = 
                                person.Schedule[scheduleDay].AllWorkTime;
                        }

                        if (person.Schedule[scheduleDay].NightWorkTime != 0)
                        {
                            wb.Workbook.Worksheets[fairSheetName].Cells[row + 3, column].Value = "Я";
                            wb.Workbook.Worksheets[fairSheetName].Cells[row, column].Value =
                                person.Schedule[scheduleDay].NightWorkTime;
                            wb.Workbook.Worksheets[fairSheetName].Cells[row + 2, column].Value += "Н";
                        }

                        if (person.Schedule[scheduleDay].Truancy)
                        {
                            wb.Workbook.Worksheets[fairSheetName].Cells[row + 3, column].Value += "НН";
                        }

                        if (person.Schedule[scheduleDay].EducationalLeave)
                        {
                            wb.Workbook.Worksheets[fairSheetName].Cells[row + 3, column].Value += "У";
                        }

                        if (person.Schedule[scheduleDay].VacationDay)
                        {
                            wb.Workbook.Worksheets[fairSheetName].Cells[row + 3, column].Value += "ОТ";
                        }

                        if (person.Schedule[scheduleDay].UnpaidLeave)
                        {
                            wb.Workbook.Worksheets[fairSheetName].Cells[row + 3, column].Value += "ДО";
                        }

                        if (person.Schedule[scheduleDay].SickDay)
                        {
                            wb.Workbook.Worksheets[fairSheetName].Cells[row + 3, column].Value += "Б";
                        }

                        if (person.Schedule[scheduleDay].DayOff)
                        {
                            if (wb.Workbook.Worksheets[fairSheetName].Cells[row + 3, column].Value == null)
                            {
                                wb.Workbook.Worksheets[fairSheetName].Cells[row + 3, column].Value =
                                    wb.Workbook.Worksheets[fairSheetName].Cells[row + 3, column].Value?.ToString() + "В";
                            }
                        }

                        scheduleDay++;
                    }

                    row += 4;
                    id++;
                }

                wb.SaveAs(new FileInfo(outputName));
            }
        }
    }
}
