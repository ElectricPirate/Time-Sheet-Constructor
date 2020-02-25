using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;



namespace Time_Sheet_Constructor.Model
{
    /// <summary>
    /// Парсер отчества и табельных номеров
    /// </summary>
    public class EmpoyeeIDParser
    {    
        ///// <summary>
        ///// Путь к .xls файлу со списком сотрудников
        ///// </summary>
        public string EmployeeFilePath_xls { get; set; }

        /// <summary>
        /// Путь к .xlsx файлу со списком сотрудников
        /// </summary>
        private string employeeFilePath_xlsx;

        List<Person> persons;

        /// <summary>
        /// Данные файла
        /// </summary>        
        private ExcelPackage excel;

        private FileInfo FI;

        /// <summary>
        /// Номер первой строки с ФИО
        /// </summary>
        private int firstFIORow;

        /// <summary>
        /// Номер столбца с ФИО
        /// </summary>
        private int fioColumn;

        /// <summary>
        /// Номер последней строки с ФИО
        /// </summary>
        private int lastFIORow;

        /// <summary>
        /// Адрес ячейки заголовка столбца списка сотрудников
        /// </summary>
        private (int row, int column) employeeCellAddress;
        private string employeeCellText = "Сотрудник";

        /// <summary>
        /// Номер первой даты приема
        /// </summary>
        private int firstDateOfReceiptRow;

        /// <summary>
        /// Номер столбца с датами приема
        /// </summary>
        private int dateOfReceiptColumn;

        /// <summary>
        /// Номер последней строки даты приема
        /// </summary>
        private int lastDateOfReceiptRow;

        /// <summary>
        /// Адрес ячейки заголовка столбца даты приема
        /// </summary>
        private (int row, int column) dateOfReceiptCellAddress;
        private string dateOfReceiptCellText = "Дата приема";

        public EmpoyeeIDParser(List<Person> persons, string _employeeFilePath_xls)
        {
            EmployeeFilePath_xls = _employeeFilePath_xls;
            this.persons = persons;
            FI = GetFI(EmployeeFilePath_xls);
            excel = new ExcelPackage(FI);
            employeeCellAddress = GetCellAddress(employeeCellText);
            dateOfReceiptCellAddress = GetCellAddress(dateOfReceiptCellText);
            firstFIORow = employeeCellAddress.row + 1;
            fioColumn = employeeCellAddress.column;
            lastFIORow = GetLastFIORow();
            firstDateOfReceiptRow = dateOfReceiptCellAddress.row + 1;
            dateOfReceiptColumn = dateOfReceiptCellAddress.column;
            lastDateOfReceiptRow = lastFIORow;
        }

        /// <summary>
        /// Конвертер .xls в .xlsx
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        private string ConvertXLS_XLSX(FileInfo file)
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

        private FileInfo GetFI(string path)
        {
            FileInfo Fi_xls = new FileInfo(path);

            if (Fi_xls.Extension == ".xls")
            {
                employeeFilePath_xlsx = ConvertXLS_XLSX(Fi_xls);
            }
            else
            {
                employeeFilePath_xlsx = path;
            }

            return new FileInfo(employeeFilePath_xlsx);
        }

        /// <summary>
        /// Добавление отчества и табельного номера
        /// </summary>
        /// <param name="persons"></param>
        /// <returns></returns>
        public List<Person> Parse()
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
                        person.DateOfReceipt = personWithId.DateOfReceipt;
                    }
                }
            }
            
            return persons;
        }

        /// <summary>
        /// Получение ФИО и табельных номеров
        /// </summary>
        /// <returns></returns>
        private List<Person> GetPersons()
        {
            var persons = new List<Person>();

            var column = fioColumn;

            var sheet = excel.Workbook.Worksheets[1];
            
                for (var row = firstFIORow; row <= lastFIORow; row++)
                {
                    var names = sheet.Cells[row, column].Value.ToString().Split(' ');
                    var currentId = Convert.ToInt32(sheet.Cells[row, column + 1].Value);
                    var currentDateOfReceipt = Convert.ToDateTime(sheet.Cells[row, column - 15].Value);
                    var currentPerson = new Person {LastName = names[0], FirstName = names[1], MiddleName = names[2], EmployeeId = currentId, DateOfReceipt=currentDateOfReceipt};
                    persons.Add(currentPerson);
                }            

            return persons;
        }

        /// <summary>
        /// Получение номера последней строки с ФИО
        /// </summary>
        /// <returns></returns>
        private int GetLastFIORow()
        {
            return excel.Workbook.Worksheets[1].Dimension.End.Row;
        }

        /// <summary>
        /// Получение номера строки ячейки "Сотрудник"
        /// </summary>
        /// <returns></returns>
        private (int,int) GetCellAddress(string searchword)
        {            
            var address = (row: 1, column: 1);
            var sheet = excel.Workbook.Worksheets[1];
            
                var query = sheet.Cells[1, 1, sheet.Dimension.End.Row, sheet.Dimension.End.Column]
                    .Where(cell => cell.Value?.ToString() == searchword);

                foreach (var cell in query)
                {
                    address.row = cell.Start.Row;
                    address.column = cell.Start.Column;                    
                    break;
                }            

            return address;
        }                
    }
}
