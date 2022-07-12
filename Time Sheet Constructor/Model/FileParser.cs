using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Windows;

namespace Time_Sheet_Constructor.Model
{
    /// <summary>
    /// Парсер данных из отчета
    /// </summary>
    public class FileParser 
    {
        /// <summary>
        /// Количество дней в месяце
        /// </summary>
        int daysCount;

        /// <summary>
        /// Первая дата из отчета, для определения месяца выгрузки
        /// </summary>
        public DateTime FirstTableDate { get; set; }

        ExcelPackage file;

        /// <summary>
        /// Список операторов
        /// </summary>
        List<Person> persons;       

        public FileParser(ExcelPackage excelReport)
        {
            file = excelReport;
            persons = GetPersons();            
            daysCount = 31;
            FirstTableDate = GetFirstTableDate();
        }

        /// <summary>
        /// Парсер
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        public List<Person> GetData()
        {     
            GetAllWorkTime();
            GetNightWorkTime();
            GetOverTimes();
            GetSickDays();
            GetVacationDays();
            GetUnpaidLeaves();
            GetEducationalLeaves();
            GetTruancys();
            GetHookies();
            GetMaternityes();
            GetPaidDaysOff();
            GetDaysOff();

            return persons;

        }
        
        /// <summary>
        /// Получаем дату
        /// </summary>
        /// <returns></returns>
        private DateTime GetFirstTableDate()
        {
            const string sheetName = "Всего";
            var firstdaterow = GetPersonCellRow(sheetName) - 1;
            var firstdatecolumn = 2;
            DateTime date;

            var d = file.Workbook.Worksheets[sheetName].Cells[firstdaterow, firstdatecolumn].Value; 

            if (d != null && DateTime.TryParse(d.ToString(), out date))
            {
                return date;
            }
            else
            {
                MessageBox.Show("Не могу понять, за какой месяц табель. \nПроверка даты оформления будет отключена.", "Внимание");
                return default;
            }            
        }

        /// <summary>
        /// Вычисление строки ячейки "Person"
        /// </summary>
        /// <param name="file"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        private int GetPersonCellRow(string sheetName)
        {
            const string searchWord = "Person";
            var getPersonalCellRow = 0;

            var sheet = file.Workbook.Worksheets[sheetName];
            
                for (var rowIndex = 1; rowIndex < 50; rowIndex++)
                {
                    if (sheet.Cells[rowIndex, 1].Value == null)
                    {
                        continue;
                    }

                    if (sheet.Cells[rowIndex, 1].Value.Equals(searchWord))
                    {
                        getPersonalCellRow = rowIndex;
                        break;
                    }
                }
            
            return getPersonalCellRow;
        }


        //private int GetPersonCellRow(string sheetName)
        //{
        //    const string searchWord = "Person";
        //    var getPersonalCellRow = 0;

        //    using (var sheet = file.Workbook.Worksheets[sheetName])
        //    {
        //        for (var rowIndex = 1; rowIndex < 50; rowIndex++)
        //        {
        //            if (sheet.Cells[rowIndex, 1].Value == null)
        //            {
        //                continue;
        //            }

        //            if (sheet.Cells[rowIndex, 1].Value.Equals(searchWord))
        //            {
        //                getPersonalCellRow = rowIndex;
        //                break;
        //            }
        //        }
        //    }

        //    return getPersonalCellRow;
        //}

        /// <summary>
        /// Последеняя строка с фио
        /// </summary>
        /// <param name="file"></param>
        /// <param name="sheetName"></param>        
        private int GetLastRowNumber(string sheetName)
        {
            var currentRow = GetPersonCellRow(sheetName);

            var sheet = file.Workbook.Worksheets[sheetName];
            
                while (!sheet.Cells[currentRow, 1].Value.Equals(null))
                {
                    currentRow++;

                    if (sheet.Cells[currentRow + 1, 1].Value == null)
                    {
                        break;
                    }
                }            

            return currentRow;
        }

        /// <summary>
        /// Получение ФИО
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        private List<Person> GetPersons()
        {
            const string sheetName = "Всего";
            var persons = new List<Person>();
            var firstFioLine = GetPersonCellRow(sheetName) + 1;
            var lastLineFio = GetLastRowNumber(sheetName);

            var sheet = file.Workbook.Worksheets[sheetName];
                           
                for (var row = firstFioLine; row <= lastLineFio; row++)
                {
                    var names = sheet.Cells[row, 1].Value.ToString().Split(' ');
                    var lastName = names[0];
                    var firstName = names[1];
                    persons.Add(new Person() { FirstName = firstName, LastName = lastName });
                }            

            return persons;
        }

        /// <summary>
        /// Получение всех рабочих часов
        /// </summary>
        /// <param name="file"></param>
        /// <param name="persons"></param>
        /// <returns></returns>
        private void GetAllWorkTime()
        {
            const string sheet = "Всего";
            var firstFioLine = GetPersonCellRow(sheet) + 1;
            var lastLineFio = GetLastRowNumber(sheet);
            var firstDayColumn = 2;
            var lastDayColumn = daysCount + 1;

            foreach (var person in persons)
            {
                for (var row = firstFioLine; row <= lastLineFio; row++)
                {
                    if (person.GetShortName().Equals(file.Workbook.Worksheets[sheet].Cells[row, 1].Value))
                    {
                        for (var column = firstDayColumn; column <= lastDayColumn; column++)
                        {
                            var dayNumber = column - 1;
                            var current = file.Workbook.Worksheets[sheet].Cells[row, column].Value;

                            if (current == null)
                            {
                                person.Schedule.Add(new Day { AllWorkTime = 0, Number = dayNumber });
                            }
                            else
                            {
                                person.Schedule.Add(new Day { AllWorkTime = Convert.ToDouble(current), Number = dayNumber });
                            }

                            dayNumber++;
                        }
                    }
                }
            }            
        }

        /// <summary>
        /// Получение сверхурочных часов
        /// </summary>
        /// <param name="file"></param>
        /// <param name="persons"></param>
        /// <returns></returns>
        private void GetOverTimes()
        {
            const string sheet = "Овертаймы";
            var firstFioLine = GetPersonCellRow(sheet) + 1;
            var lastLineFio = GetLastRowNumber(sheet);
            var firstDayColumn = 2;
            var lastDayColumn = daysCount + 1;

            foreach (var person in persons)
            {
                for (var row = firstFioLine; row <= lastLineFio; row++)
                {
                    if (person.GetShortName().Equals(file.Workbook.Worksheets[sheet].Cells[row, 1].Value))
                    {
                        for (var column = firstDayColumn; column <= lastDayColumn; column++)
                        {
                            var dayIndex = column - 2;
                            var current = file.Workbook.Worksheets[sheet].Cells[row, column].Value;

                            if (current == null)
                            {
                                person.Schedule[dayIndex].OverTime = 0;
                            }
                            else
                            {
                                person.Schedule[dayIndex].OverTime = Convert.ToDouble(current);
                            }

                            dayIndex++;
                        }
                    }
                }
            }            
        }

        /// <summary>
        /// Получение ночных часов
        /// </summary>
        /// <param name="file"></param>
        /// <param name="persons"></param>
        /// <returns></returns>
        private void GetNightWorkTime()
        {
            const string sheet = "Ночные";
            var firstFioLine = GetPersonCellRow(sheet) + 1;
            var lastLineFio = GetLastRowNumber(sheet);
            var firstDayColumn = 2;
            var lastDayColumn = daysCount + 1;

            foreach (var person in persons)
            {
                for (var row = firstFioLine; row <= lastLineFio; row++)
                {
                    if (person.GetShortName().Equals(file.Workbook.Worksheets[sheet].Cells[row, 1].Value))
                    {
                        for (var column = firstDayColumn; column <= lastDayColumn; column++)
                        {
                            var dayIndex = column - 2;
                            var current = file.Workbook.Worksheets[sheet].Cells[row, column].Value;

                            if (current == null)
                            {
                                person.Schedule[dayIndex].NightWorkTime = 0;
                            }
                            else
                            {
                                person.Schedule[dayIndex].NightWorkTime = Convert.ToDouble(current);
                            }

                            dayIndex++;
                        }
                    }
                }
            }            
        }                       

        /// <summary>
        /// Получение больничных
        /// </summary>
        /// <param name="file"></param>
        /// <param name="persons"></param>
        /// <returns></returns>
        private void GetSickDays()
        {
            const string sheet = "Больничные";
            var firstFioLine = GetPersonCellRow(sheet) + 1;
            var lastLineFio = GetLastRowNumber(sheet);
            var firstDayColumn = 2;
            var lastDayColumn = daysCount + 1;

            foreach (var person in persons)
            {
                for (var row = firstFioLine; row <= lastLineFio; row++)
                {
                    if (person.GetShortName().Equals(file.Workbook.Worksheets[sheet].Cells[row, 1].Value))
                    {
                        for (var column = firstDayColumn; column <= lastDayColumn; column++)
                        {
                            var dayIndex = column - 2;
                            var current = file.Workbook.Worksheets[sheet].Cells[row, column].Value;

                            if (current != null)
                            {
                                person.Schedule[dayIndex].SickDay = current.ToString();
                            }

                            dayIndex++;
                        }
                    }
                }
            }            
        }

        /// <summary>
        /// Получение ежегодных отпусков
        /// </summary>
        /// <param name="file"></param>
        /// <param name="persons"></param>
        /// <returns></returns>
        private void GetVacationDays()
        {
            const string sheet = "Отпуск";
            var firstFioLine = GetPersonCellRow(sheet) + 1;
            var lastLineFio = GetLastRowNumber(sheet);
            var firstDayColumn = 2;
            var lastDayColumn = daysCount + 1;

            foreach (var person in persons)
            {
                for (var row = firstFioLine; row <= lastLineFio; row++)
                {
                    if (person.GetShortName().Equals(file.Workbook.Worksheets[sheet].Cells[row, 1].Value))
                    {
                        for (var column = firstDayColumn; column <= lastDayColumn; column++)
                        {
                            var dayIndex = column - 2;
                            var current = file.Workbook.Worksheets[sheet].Cells[row, column].Value;

                            if (current != null)
                            {
                                person.Schedule[dayIndex].VacationDay = current.ToString();
                            }
                            
                            dayIndex++;
                        }
                    }
                }
            }            
        }

        /// <summary>
        /// Получение дополнительных отпусков
        /// </summary>
        /// <param name="file"></param>
        /// <param name="persons"></param>
        /// <returns></returns>
        private void GetUnpaidLeaves()
        {
            const string sheet = "Неоп_Отпуск";
            var firstFioLine = GetPersonCellRow(sheet) + 1;
            var lastLineFio = GetLastRowNumber(sheet);
            var firstDayColumn = 2;
            var lastDayColumn = daysCount + 1;

            foreach (var person in persons)
            {
                for (var row = firstFioLine; row <= lastLineFio; row++)
                {
                    if (person.GetShortName().Equals(file.Workbook.Worksheets[sheet].Cells[row, 1].Value))
                    {
                        for (var column = firstDayColumn; column <= lastDayColumn; column++)
                        {
                            var dayIndex = column - 2;
                            var current = file.Workbook.Worksheets[sheet].Cells[row, column].Value;

                            if (current != null)
                            {
                                person.Schedule[dayIndex].UnpaidLeave = current.ToString();
                            }

                            dayIndex++;
                        }
                    }
                }
            }            
        }

        /// <summary>
        /// Получение учебных отпусков
        /// </summary>
        /// <param name="file"></param>
        /// <param name="persons"></param>
        /// <returns></returns>
        private void GetEducationalLeaves()
        {
            const string sheet = "Учен_Отпуск";
            var firstFioLine = GetPersonCellRow(sheet) + 1;
            var lastLineFio = GetLastRowNumber(sheet);
            var firstDayColumn = 2;
            var lastDayColumn = daysCount + 1;

            foreach (var person in persons)
            {
                for (var row = firstFioLine; row <= lastLineFio; row++)
                {
                    if (person.GetShortName().Equals(file.Workbook.Worksheets[sheet].Cells[row, 1].Value))
                    {
                        for (var column = firstDayColumn; column <= lastDayColumn; column++)
                        {
                            var dayIndex = column - 2;
                            var current = file.Workbook.Worksheets[sheet].Cells[row, column].Value;

                            if (current != null)
                            {
                                person.Schedule[dayIndex].EducationalLeave = current.ToString();
                            }

                            dayIndex++;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Получение неявок
        /// </summary>
        /// <param name="file"></param>
        /// <param name="persons"></param>
        /// <returns></returns>
        private void GetTruancys()
        {
            const string sheet = "Неявка";
            var firstFioLine = GetPersonCellRow(sheet) + 1;
            var lastLineFio = GetLastRowNumber(sheet);
            var firstDayColumn = 2;
            var lastDayColumn = daysCount + 1;

            foreach (var person in persons)
            {
                for (var row = firstFioLine; row <= lastLineFio; row++)
                {
                    if (person.GetShortName().Equals(file.Workbook.Worksheets[sheet].Cells[row, 1].Value))
                    {
                        for (var column = firstDayColumn; column <= lastDayColumn; column++)
                        {
                            var dayIndex = column - 2;
                            var current = file.Workbook.Worksheets[sheet].Cells[row, column].Value;

                            if (current != null)
                            {
                                person.Schedule[dayIndex].Truancy = current.ToString();
                            }
                            
                            dayIndex++;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Получение прогулов
        /// </summary>
        /// <param name="file"></param>
        /// <param name="persons"></param>
        /// <returns></returns>
        private void GetHookies()
        {
            const string sheet = "Прогул";
            var firstFioLine = GetPersonCellRow(sheet) + 1;
            var lastLineFio = GetLastRowNumber(sheet);
            var firstDayColumn = 2;
            var lastDayColumn = daysCount + 1;

            foreach (var person in persons)
            {
                for (var row = firstFioLine; row <= lastLineFio; row++)
                {
                    if (person.GetShortName().Equals(file.Workbook.Worksheets[sheet].Cells[row, 1].Value))
                    {
                        for (var column = firstDayColumn; column <= lastDayColumn; column++)
                        {
                            var dayIndex = column - 2;
                            var current = file.Workbook.Worksheets[sheet].Cells[row, column].Value;

                            if (current != null)
                            {
                                person.Schedule[dayIndex].Hooky = current.ToString();
                            }                            

                            dayIndex++;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Получение выходных
        /// </summary>
        /// <param name="file"></param>
        /// <param name="persons"></param>
        /// <returns></returns>
        private void GetDaysOff()
        {
            const string sheet = "Выходные";
            var firstFioLine = GetPersonCellRow(sheet) + 1;
            var lastLineFio = GetLastRowNumber(sheet);
            var firstDayColumn = 2;
            var lastDayColumn = daysCount + 1;

            foreach (var person in persons)
            {
                for (var row = firstFioLine; row <= lastLineFio; row++)
                {
                    if (person.GetShortName().Equals(file.Workbook.Worksheets[sheet].Cells[row, 1].Value))
                    {
                        for (var column = firstDayColumn; column <= lastDayColumn; column++)
                        {
                            var dayIndex = column - 2;
                            var current = file.Workbook.Worksheets[sheet].Cells[row, column].Value;

                            if (current == null)
                            {
                                person.Schedule[dayIndex].DayOff = false;
                            }
                            else
                            {
                                person.Schedule[dayIndex].DayOff = true;
                            }

                            dayIndex++;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Получение оплачиваемых выходных
        /// </summary>
        private void GetPaidDaysOff()
        {
            const string sheet = "Оплачиваемый выходной";
            var firstFioLine = GetPersonCellRow(sheet) + 1;
            var lastLineFio = GetLastRowNumber(sheet);
            var firstDayColumn = 2;
            var lastDayColumn = daysCount + 1;

            foreach (var person in persons)
            {
                for (var row = firstFioLine; row <= lastLineFio; row++)
                {
                    if (person.GetShortName().Equals(file.Workbook.Worksheets[sheet].Cells[row, 1].Value))
                    {
                        for (var column = firstDayColumn; column <= lastDayColumn; column++)
                        {
                            var dayIndex = column - 2;
                            var current = file.Workbook.Worksheets[sheet].Cells[row, column].Value;

                            if (current != null)
                            {
                                person.Schedule[dayIndex].PaidDayOff = current.ToString();
                            }
                            
                            dayIndex++;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Получение отсутствий по беременности и родам
        /// </summary>
        /// <param name="file"></param>
        /// <param name="persons"></param>
        /// <returns></returns>
        private void GetMaternityes()
        {
            const string sheet = "Декрет";
            var firstFioLine = GetPersonCellRow(sheet) + 1;
            var lastLineFio = GetLastRowNumber(sheet);
            var firstDayColumn = 2;
            var lastDayColumn = daysCount + 1;

            foreach (var person in persons)
            {
                for (var row = firstFioLine; row <= lastLineFio; row++)
                {
                    if (person.GetShortName().Equals(file.Workbook.Worksheets[sheet].Cells[row, 1].Value))
                    {
                        for (var column = firstDayColumn; column <= lastDayColumn; column++)
                        {
                            var dayIndex = column - 2;
                            var current = file.Workbook.Worksheets[sheet].Cells[row, column].Value;

                            if (current != null)
                            {
                                person.Schedule[dayIndex].MaternityLeave = current.ToString();
                            }                            

                            dayIndex++;
                        }
                    }
                }
            }
        }

    }
}
