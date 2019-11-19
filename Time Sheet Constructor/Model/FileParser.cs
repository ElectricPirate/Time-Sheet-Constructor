using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using Time_Sheet_Constructor.Model;

namespace Time_Sheet_Constructor.Model
{
    
    public static class FileParser
    {
        public static List<string> GetSheetsList(ExcelPackage file)
        {
            var sheets = new List<string>();
            
            foreach (var sheet in file.Workbook.Worksheets)
            {
                sheets.Add(sheet.Name);
            }

            return sheets;
        }

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
                persons.Add(new Person() {FirstName = firstName, LastName = lastName});
            }

            return persons;
        }


    }
}
