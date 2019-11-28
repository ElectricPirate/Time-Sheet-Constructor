using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;



namespace Time_Sheet_Constructor.Model
{
    /// <summary>
    /// Парсер отчества и табельных номеров
    /// </summary>
    public static class EmpoyeeIDParser
    {
        /// <summary>
        /// Путь к .xls файлу со списком сотрудников
        /// </summary>
        /// 
        static string EmployeeFilePath_xls =
            @"C:\Users\vadim.turetskiy\Documents\Табель\Time sheet constructor\Список от 15.11.xls";

        /// <summary>
        /// Данные файла
        /// </summary>
        static FileInfo Fi_xls => new FileInfo(EmployeeFilePath_xls);

        /// <summary>
        /// Путь к .xlsx файлу со списком сотрудников
        /// </summary>
        static string EmployeeFilePath_xlsx = ConvertXLS_XLSX(Fi_xls);

        /// <summary>
        /// Данные файла
        /// </summary>
        static FileInfo Fi_xlsx => new FileInfo(EmployeeFilePath_xlsx);


        static ExcelPackage Excel => new ExcelPackage(Fi_xlsx);

        /// <summary>
        /// Номер первой строки с ФИО
        /// </summary>
        static int FirstFIORow => GetEmployeeCellRow() + 1;

        /// <summary>
        /// Номер столбца с ФИО
        /// </summary>
        static int FirstFIOColumn => GetEmployeeCellColumn();

        /// <summary>
        /// Номер последней строки с ФИО
        /// </summary>
        static int LastFIORow => GetLastFIORow();

        /// <summary>
        /// Конвертер .xls в .xlsx
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        private static string ConvertXLS_XLSX(FileInfo file)
        {
            var app = new Microsoft.Office.Interop.Excel.Application();
            var xlsFile = file.FullName;
            var wb = app.Workbooks.Open(xlsFile);
            var xlsxFile = xlsFile + "x";
            wb.SaveAs(Filename: xlsxFile, FileFormat: Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
            wb.Close();
            app.Quit();
            return xlsxFile;
        }

        /// <summary>
        /// Добавление отчества и табельного номера
        /// </summary>
        /// <param name="persons"></param>
        /// <returns></returns>
        public static List<Person> Parse(List<Person> persons)
        {
            var personsWithId = GetPersons();

            foreach (var person in persons)
            {
                foreach (var personWithId in personsWithId)
                {
                    if (person.FirstName.Equals(personWithId.FirstName) &&
                        person.LastName.Equals(personWithId.LastName))
                    {
                        person.MiddleName = personWithId.MiddleName;
                        person.EmployeeId = personWithId.EmployeeId;
                    }
                }
            }
            
            return persons;
        }

        /// <summary>
        /// Получение ФИО и табельных номеров
        /// </summary>
        /// <returns></returns>
        private static List<Person> GetPersons()
        {
            var persons = new List<Person>();

            var column = FirstFIOColumn;
            var firstRow = FirstFIORow;
            var lastRow = LastFIORow;

            using (var sheet = Excel.Workbook.Worksheets[1])
            {
                for (var row = firstRow; row <= lastRow; row++)
                {
                    var names = sheet.Cells[row, column].Value.ToString().Split(' ');
                    var currentId = Convert.ToInt32(sheet.Cells[row, column + 1].Value);
                    var currentPerson = new Person {LastName = names[0], FirstName = names[1], MiddleName = names[2], EmployeeId = currentId};
                    persons.Add(currentPerson);
                }
            }

            return persons;
        }

        /// <summary>
        /// Получение номера последней строки с ФИО
        /// </summary>
        /// <returns></returns>
        private static int GetLastFIORow()
        {
            return Excel.Workbook.Worksheets[1].Dimension.End.Row;
        }


        /// <summary>
        /// Получение номера строки ячейки "Сотрудник"
        /// </summary>
        /// <returns></returns>
        private static int GetEmployeeCellRow()
        {
            var searchword = "Сотрудник";
            var address = 1;

            using (var sheet = Excel.Workbook.Worksheets[1])
            {
                var query = sheet.Cells[1, 1, sheet.Dimension.End.Row, sheet.Dimension.End.Column]
                    .Where(cell => cell.Value?.ToString() == searchword);

                foreach (var cell in query)
                {
                    address = cell.Start.Row;
                    break;
                }
            }

            return address;
        }

        /// <summary>
        /// Получение номера столбца ячейки "Сотрудник"
        /// </summary>
        /// <returns></returns>
        private static int GetEmployeeCellColumn()
        {
            var searchword = "Сотрудник";
            var address = 1;

            using (var sheet = Excel.Workbook.Worksheets[1])
            {
                var query = sheet.Cells[1, 1, sheet.Dimension.End.Row, sheet.Dimension.End.Column]
                    .Where(cell => cell.Value?.ToString() == searchword);

                foreach (var cell in query)
                {
                    address = cell.Start.Column;
                    break;
                }
            }

            return address;
        }
    }
}
