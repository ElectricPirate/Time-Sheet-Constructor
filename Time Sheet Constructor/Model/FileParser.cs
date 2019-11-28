using OfficeOpenXml;
using System;
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;

namespace Time_Sheet_Constructor.Model
{
    /// <summary>
    /// Парсер данных из отчета
    /// </summary>
    public static class FileParser
    {
        /// <summary>
        /// Количество дней в текущем месяце
        /// </summary>
        public static int DaysCount => DateTime.DaysInMonth(DateTime.Today.Year, DateTime.Today.Month);

        /// <summary>
        /// Парсер
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        public static List<Person> GetData(ExcelPackage file)
        {
            var persons = GetPersons(file);
            GetAllWorkTime(file, persons);
            GetNightWorkTime(file, persons);
            GetSickDays(file, persons);
            GetVacationDays(file, persons);
            GetUnpaidLeaves(file, persons);
            GetEducationalLeaves(file, persons);
            GetTruancys(file, persons);
            GetDaysOff(file, persons);

            return persons;

        }

        /// <summary>
        /// Получение списка листов
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        public static List<string> GetSheetsList(ExcelPackage file)
        {
            var sheets = new List<string>();

            foreach (var sheet in file.Workbook.Worksheets)
            {
                sheets.Add(sheet.Name);
            }

            return sheets;
        }

        /// <summary>
        /// Вычисление строки ячейки "Person"
        /// </summary>
        /// <param name="file"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public static int GetPersonCellRow(ExcelPackage file, string sheetName)
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

        /// <summary>
        /// Последеняя строка с фио
        /// </summary>
        /// <param name="file"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public static int GetLastRowNumber(ExcelPackage file, string sheetName)
        {
            var currentRow = GetPersonCellRow(file, sheetName);

            while (!file.Workbook.Worksheets[sheetName].Cells[currentRow, 1].Value.Equals(null))
            {
                currentRow++;

                if (file.Workbook.Worksheets[sheetName].Cells[currentRow + 1, 1].Value == null)
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
        public static List<Person> GetPersons(ExcelPackage file)
        {
            const string sheet = "Всего";
            var persons = new List<Person>();

            var firstFioLine = GetPersonCellRow(file, sheet) + 1;
            var lastLineFio = GetLastRowNumber(file, sheet);

            for (var row = firstFioLine; row <= lastLineFio; row++)
            {
                var names = file.Workbook.Worksheets[sheet].Cells[row, 1].Value.ToString().Split(' ');
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
        public static List<Person> GetAllWorkTime(ExcelPackage file, List<Person> persons)
        {
            const string sheet = "Всего";
            var firstFioLine = GetPersonCellRow(file, sheet) + 1;
            var lastLineFio = GetLastRowNumber(file, sheet);
            var firstDayColumn = 2;
            var lastDayColumn = DaysCount + 1;

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

            return persons;
        }

        /// <summary>
        /// Получение ночных часов
        /// </summary>
        /// <param name="file"></param>
        /// <param name="persons"></param>
        /// <returns></returns>
        public static List<Person> GetNightWorkTime(ExcelPackage file, List<Person> persons)
        {
            const string sheet = "Ночные";
            var firstFioLine = GetPersonCellRow(file, sheet) + 1;
            var lastLineFio = GetLastRowNumber(file, sheet);
            var firstDayColumn = 2;
            var lastDayColumn = DaysCount + 1;

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

            return persons;
        }


        /// TODO: оптимизировать методы ниже, они ужасные

        /// <summary>
        /// Получение больничных
        /// </summary>
        /// <param name="file"></param>
        /// <param name="persons"></param>
        /// <returns></returns>
        private static List<Person> GetSickDays(ExcelPackage file, List<Person> persons)
        {
            const string sheet = "Больничные";
            var firstFioLine = GetPersonCellRow(file, sheet) + 1;
            var lastLineFio = GetLastRowNumber(file, sheet);
            var firstDayColumn = 2;
            var lastDayColumn = DaysCount + 1;

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
                                person.Schedule[dayIndex].SickDay = false;
                            }
                            else
                            {
                                person.Schedule[dayIndex].SickDay = true;
                            }

                            dayIndex++;
                        }
                    }
                }
            }

            return persons;
        }

        /// <summary>
        /// Получение ежегодных отпусков
        /// </summary>
        /// <param name="file"></param>
        /// <param name="persons"></param>
        /// <returns></returns>
        private static List<Person> GetVacationDays(ExcelPackage file, List<Person> persons)
        {
            const string sheet = "Отпуск";
            var firstFioLine = GetPersonCellRow(file, sheet) + 1;
            var lastLineFio = GetLastRowNumber(file, sheet);
            var firstDayColumn = 2;
            var lastDayColumn = DaysCount + 1;

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
                                person.Schedule[dayIndex].VacationDay = false;
                            }
                            else
                            {
                                person.Schedule[dayIndex].VacationDay = true;
                            }

                            dayIndex++;
                        }
                    }
                }
            }

            return persons;
        }

        /// <summary>
        /// Получение дополнительных отпусков
        /// </summary>
        /// <param name="file"></param>
        /// <param name="persons"></param>
        /// <returns></returns>
        private static List<Person> GetUnpaidLeaves(ExcelPackage file, List<Person> persons)
        {
            const string sheet = "Неоп_Отпуск";
            var firstFioLine = GetPersonCellRow(file, sheet) + 1;
            var lastLineFio = GetLastRowNumber(file, sheet);
            var firstDayColumn = 2;
            var lastDayColumn = DaysCount + 1;

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
                                person.Schedule[dayIndex].UnpaidLeave = false;
                            }
                            else
                            {
                                person.Schedule[dayIndex].UnpaidLeave = true;
                            }

                            dayIndex++;
                        }
                    }
                }
            }

            return persons;
        }

        /// <summary>
        /// Получение учебных отпусков
        /// </summary>
        /// <param name="file"></param>
        /// <param name="persons"></param>
        /// <returns></returns>
        private static List<Person> GetEducationalLeaves(ExcelPackage file, List<Person> persons)
        {
            const string sheet = "Учен_Отпуск";
            var firstFioLine = GetPersonCellRow(file, sheet) + 1;
            var lastLineFio = GetLastRowNumber(file, sheet);
            var firstDayColumn = 2;
            var lastDayColumn = DaysCount + 1;

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
                                person.Schedule[dayIndex].EducationalLeave = false;
                            }
                            else
                            {
                                person.Schedule[dayIndex].EducationalLeave = true;
                            }

                            dayIndex++;
                        }
                    }
                }
            }

            return persons;
        }

        /// <summary>
        /// Получение неявок
        /// </summary>
        /// <param name="file"></param>
        /// <param name="persons"></param>
        /// <returns></returns>
        private static List<Person> GetTruancys(ExcelPackage file, List<Person> persons)
        {
            const string sheet = "Прогул";
            var firstFioLine = GetPersonCellRow(file, sheet) + 1;
            var lastLineFio = GetLastRowNumber(file, sheet);
            var firstDayColumn = 2;
            var lastDayColumn = DaysCount + 1;

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
                                person.Schedule[dayIndex].Truancy = false;
                            }
                            else
                            {
                                person.Schedule[dayIndex].Truancy = true;
                            }

                            dayIndex++;
                        }
                    }
                }
            }

            return persons;
        }

        /// <summary>
        /// Получение выходных
        /// </summary>
        /// <param name="file"></param>
        /// <param name="persons"></param>
        /// <returns></returns>
        private static List<Person> GetDaysOff(ExcelPackage file, List<Person> persons)
        {
            const string sheet = "Выходные";
            var firstFioLine = GetPersonCellRow(file, sheet) + 1;
            var lastLineFio = GetLastRowNumber(file, sheet);
            var firstDayColumn = 2;
            var lastDayColumn = DaysCount + 1;

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

            return persons;
        }

    }
}
