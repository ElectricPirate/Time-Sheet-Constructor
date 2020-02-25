using System;
using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using Microsoft.Win32;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.FormulaParsing.Utilities;
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
            DataContext = new ViewModel(new DefaultDialogService());
        }

        private void About_Click(object sender, RoutedEventArgs e)
        {
            var AboutWindow = new AboutWindow();
            AboutWindow.Show();
        }
    }
}
