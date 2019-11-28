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
    public static class EmpoyeeIDParser
    {
        static string EmployeeFilePath =
            @"C:\Users\vadim.turetskiy\Documents\Табель\Time sheet constructor\Список от 15.11.xlsx";

        static FileInfo Fi => new FileInfo(EmployeeFilePath);
        static ExcelPackage Excel => new ExcelPackage(Fi);

        static int FirstFIORow => GetEmployeeCellRow() + 1;
        static int FirstFIOColumn => GetEmployeeCellColumn();
        static int LastFIORow => GetLastFIORow();
        static int FirstIDRow => FirstFIORow;
        static int FirstIDColumn => FirstFIOColumn + 1;
        static int LastIDRow => LastFIORow;

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

        private static int GetLastFIORow()
        {
            return Excel.Workbook.Worksheets[1].Dimension.End.Row;
        }

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
