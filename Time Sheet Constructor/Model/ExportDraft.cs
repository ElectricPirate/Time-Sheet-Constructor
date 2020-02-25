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
        string tableLayoutPath;
       
        /// <summary>
        /// Путь выходного файла табеля
        /// </summary>
        string outputName;            

        /// <summary>
        /// Имя листа Черновик
        /// </summary>
        const string draftSheetName = "Черновик";

        /// <summary>
        /// Номер строки первого ФИО
        /// </summary>
        int firstFioRow;

        /// <summary>
        /// Номер столбца с ФИО
        /// </summary>
        int fioColumn;

        /// <summary>
        /// Номер столбца с табельными номерами
        /// </summary>
        int emloyeeIdColumn;
        
        /// <summary>
        /// Начальный столбец
        /// </summary>
        int firstDay;

        /// <summary>
        /// Конечный столбец
        /// </summary>        
        int lastDay;

        /// <summary>
        /// Количество человек, у которых первый рабочий день раньше даты приема
        /// </summary>
        int firstWorkDayErrors;

        /// <summary>
        /// Количество дней, у которых рабочее время пересекается с отсутствием
        /// </summary>
        int worktimeCrossingErrors;

        /// <summary>
        /// Данные файла шаблона
        /// </summary>
        FileInfo fi;

        ExcelPackage excel;

        List<Person> persons;

        public ExportDraft(string tableLayoutPath, List<Person> persons,int firstDay,int lastDay)
        {
            this.tableLayoutPath = tableLayoutPath;
            fi = new FileInfo(tableLayoutPath);
            excel = new ExcelPackage(fi);
            fioColumn = 1;
            firstFioRow = 2;
            emloyeeIdColumn = fioColumn + 1;
            this.firstDay = firstDay + 2;
            this.lastDay = lastDay + 2;            
            this.persons = persons;
            outputName = $"{fi.DirectoryName}\\Табель выход.xlsx";
        }

        /// <summary>
        /// Пишем данные в файл
        /// </summary>
        /// <param name="persons"></param>
        public void Write()
        {
            using (var wb = excel)
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

                    if (person.FirstWorkDay < person.DateOfReceipt.Day && person.DateOfReceipt.Year == DateTime.Now.Year && person.DateOfReceipt.Month == DateTime.Now.Month)
                    {
                        wb.Workbook.Worksheets[draftSheetName].Cells[row, fioColumn].
                            AddComment($"Внимание! Первый рабочий день ({person.FirstWorkDay}) раньше даты приема ({person.DateOfReceipt.Day})", "Автор");
                        wb.Workbook.Worksheets[draftSheetName].Cells[row, fioColumn].Style.Font.Color.SetColor(System.Drawing.Color.Red);
                        firstWorkDayErrors++;
                    }

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

                        if (person.Schedule[scheduleDay].Crossing)
                        {
                            if (wb.Workbook.Worksheets[draftSheetName].Cells[row, fioColumn].Comment == null)
                            {
                                wb.Workbook.Worksheets[draftSheetName].Cells[row, fioColumn].
                            AddComment($"Внимание! Рабочее время пересекается с отсутствием!", "Автор");
                            }
                            else
                            {
                                wb.Workbook.Worksheets[draftSheetName].Cells[row, fioColumn].Comment.Text += "\nРабочее время пересекается с отсутствием!";                            
                            }
                            
                            wb.Workbook.Worksheets[draftSheetName].Cells[row, fioColumn].Style.Font.Color.SetColor(System.Drawing.Color.Red);
                            wb.Workbook.Worksheets[draftSheetName].Cells[row, column].Style.Font.Color.SetColor(System.Drawing.Color.Red);

                            worktimeCrossingErrors++;
                        }

                        scheduleDay++;
                    }

                    row++;
                }            

                if (firstWorkDayErrors > 0)
                {
                    MessageBox.Show($"У {firstWorkDayErrors} операторов первый рабочий день раньше оформления. \nОбращайте внимание на комментарии.", "Внимание!", MessageBoxButton.OK, MessageBoxImage.Error);
                }

                if (worktimeCrossingErrors > 0)
                {
                    MessageBox.Show($"Обнаружено {worktimeCrossingErrors} случаев пересечения рабочего времени с отсутствием, необходимо перенести часы. \nОбращайте внимание на комментарии.", "Внимание!", MessageBoxButton.OK, MessageBoxImage.Error);
                }

                try
                {
                    wb.SaveAs(new FileInfo(outputName));
                    MessageBox.Show($"Файл сохранен в {outputName}", "Успешно!");
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                }                
                
            }
        }
    }
}
