using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using Time_Sheet_Constructor.Model;

namespace Time_Sheet_Constructor
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            const string filepath = @"C:\Users\vadim.turetskiy\Documents\Табель\Time sheet constructor\ТабельСПБ.xlsx";
            var existingFile = new FileInfo(filepath);
            var table = new ExcelPackage(existingFile);
            SheetsList.ItemsSource = FileParser.GetSheetsList(table);

            var Cells = new List<string>();
            var Sheets = FileParser.GetSheetsList(table);

            foreach (var sheet in Sheets)
            {
                var currentPers = FileParser.GetPersonCellRow(table, sheet);
                var currentLastRow = FileParser.GetLastRowNumber(table, sheet);
                Cells.Add($"Person: {currentPers}, Last Row: {currentLastRow}");
            }

            SheetsCells.ItemsSource = Cells;

            var Persons = FileParser.GetPersons(table);
            var stringPers = new List<string>();
            foreach (var person in Persons)
            {
                stringPers.Add(person.ToString());
            }
            PersonsList.ItemsSource = stringPers;



        }

        
    }
}
