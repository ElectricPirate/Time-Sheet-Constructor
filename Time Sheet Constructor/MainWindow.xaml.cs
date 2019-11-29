using System;
using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
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
            var TableExport = new ExcelPackage();
            
            Persons = EmpoyeeIDParser.Parse(Persons);

            DataExport.FirstHalf(Persons);

            var Problems = new List<Tuple<string, int>>();

            foreach (var person in Persons)
            {
                if (person.EmployeeId == 0)
                {
                    continue;
                }

                foreach (var day in person.Schedule)
                {
                    if (day.IsCrossing)
                    {
                        var problem = (name: person.GetShortName(), day: day.Number);
                        Problems.Add(new Tuple<string, int> (problem.name, problem.day));
                    }
                }
            }

            MessageBox.Show($"Проблем: {Problems.Count.ToString()}");

            //foreach (var problem in Problems)
            //{
            //    MessageBox.Show($"Обнаружено пересечение: {problem.Item1} {problem.Item2.ToString()}");
            //}

            var s = 1;
            






















        }


    }
}
