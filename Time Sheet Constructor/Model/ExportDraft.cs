using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using OfficeOpenXml;
using Time_Sheet_Constructor.Annotations;

namespace Time_Sheet_Constructor.Model 
{
    /// <summary>
    /// Заполнение черновика
    /// </summary>
    public class ExportDraft
    {
        /// <summary>
        /// Путь шаблона табеля
        /// </summary>
        public static string TableLayoutPath { get; set; }
       
        /// <summary>
        /// Путь выходного файла табеля
        /// </summary>
        static string outputName => $"{Fi.DirectoryName}\\Табель выход.xlsx";            

        /// <summary>
        /// Имя листа Черновик
        /// </summary>
        const string draftSheetName = "Черновик";

        /// <summary>
        /// Номер строки первого ФИО
        /// </summary>
        static int firstFioRow = 2;

        /// <summary>
        /// Номер столбца с ФИО
        /// </summary>
        static int fioColumn = 1;

        /// <summary>
        /// Номер столбца с табельными номерами
        /// </summary>
        static int emloyeeIdColumn = fioColumn + 1;
        
        /// <summary>
        /// Начальный столбец
        /// </summary>
        private static int firstDay = ViewModel.FirstDay + 2;
        
        /// <summary>
        /// Конечный столбец
        /// </summary>
        private static int lastDay = ViewModel.LastDay + 2;

        /// <summary>
        /// Данные файла шаблона
        /// </summary>
        static FileInfo Fi => new FileInfo(TableLayoutPath);

        static ExcelPackage Excel => new ExcelPackage(Fi);

        /// <summary>
        /// Пишем данные в файл
        /// </summary>
        /// <param name="persons"></param>
        public static void Write(List<Person> persons)
        {
            using (var wb = Excel)
            {
                var row = firstFioRow;

                foreach (var person in persons)
                {
                    if (person.EmployeeId == 0)
                    {
                        continue;
                    }
                    
                    var scheduleDay = 0;

                    wb.Workbook.Worksheets[draftSheetName].Cells[row, fioColumn].Value = person.GetFullName();
                    wb.Workbook.Worksheets[draftSheetName].Cells[row, emloyeeIdColumn].Value = person.EmployeeId;
                    
                    for (var column = firstDay; column <= lastDay; column++)
                    {
                        if (person.Schedule[scheduleDay].AllWorkTime != 0)
                        {
                            if (person.Schedule[scheduleDay].NightWorkTime != 0)
                            {
                                wb.Workbook.Worksheets[draftSheetName].Cells[row, column].Value =
                                    $"{Math.Round(person.Schedule[scheduleDay].AllWorkTime, 1).ToString()}/{Math.Round(person.Schedule[scheduleDay].NightWorkTime, 1).ToString()}";
                            }
                            else
                            {
                                wb.Workbook.Worksheets[draftSheetName].Cells[row, column].Value = $"{Math.Round(person.Schedule[scheduleDay].AllWorkTime, 1).ToString()}";
                            }
                        }

                        if (person.Schedule[scheduleDay].SickDay)
                        {
                            wb.Workbook.Worksheets[draftSheetName].Cells[row, column].Value += "Б";
                        }

                        if (person.Schedule[scheduleDay].VacationDay)
                        {
                            wb.Workbook.Worksheets[draftSheetName].Cells[row, column].Value += "ОТ";
                        }

                        if (person.Schedule[scheduleDay].UnpaidLeave)
                        {
                            wb.Workbook.Worksheets[draftSheetName].Cells[row, column].Value += "ДО";
                        }

                        if (person.Schedule[scheduleDay].EducationalLeave)
                        {
                            wb.Workbook.Worksheets[draftSheetName].Cells[row, column].Value += "У";
                        }

                        if (person.Schedule[scheduleDay].Truancy)
                        {
                            wb.Workbook.Worksheets[draftSheetName].Cells[row, column].Value += "НН";
                        }

                        if (person.Schedule[scheduleDay].MaternityLeave)
                        {
                            wb.Workbook.Worksheets[draftSheetName].Cells[row, column].Value += "ОЖ";
                        }

                        if (person.Schedule[scheduleDay].DayOff)
                        {
                            if (wb.Workbook.Worksheets[draftSheetName].Cells[row, column].Value == null)
                            {
                                wb.Workbook.Worksheets[draftSheetName].Cells[row, column].Value = "В";
                            }
                        }

                        if (person.Schedule[scheduleDay].OverTime != 0 && person.Schedule[scheduleDay].AllWorkTime == 0)
                        {
                            if (wb.Workbook.Worksheets[draftSheetName].Cells[row, column].Value == null)
                            {
                                wb.Workbook.Worksheets[draftSheetName].Cells[row, column].Value = "В";
                            }
                        }

                        scheduleDay++;
                    }

                    row++;
                }
                
                wb.SaveAs(new FileInfo(outputName));
                MessageBox.Show($"Файл сохранен в {ExportDraft.outputName}", "Успешно");
            }
        }
    }
}
