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

            var Persons = FileParser.GetData(table);
            var p = 2;







        }


    }
}
