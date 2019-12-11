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
using Time_Sheet_Constructor.Model;

namespace Time_Sheet_Constructor
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public static string EmployeeFilePath { get; set; }

        public static string TeleoptiReportPath { get; set; }

        public static int FirstDay { get; set; }

        public static int LastDay { get; set; }

        public MainWindow()
        {
            InitializeComponent();
        }


        private void TeleoptiReportPath_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog myDialog = new OpenFileDialog();
            myDialog.Filter = "Книга Excel(*.xlsx;*.xls)|*.xlsx;*.xls" + "|Все файлы (*.*)|*.* ";
            myDialog.CheckFileExists = true;
            myDialog.Multiselect = true;
            if (myDialog.ShowDialog() == true)
            {
                TextBox_TeleoptiReportPath.Text = myDialog.FileName;
                TeleoptiReportPath = myDialog.FileName;
            }
        }

        private void EmployeeFilePath_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog myDialog = new OpenFileDialog();
            myDialog.Filter = "Книга Excel(*.xlsx;*.xls)|*.xlsx;*.xls" + "|Все файлы (*.*)|*.* ";
            myDialog.CheckFileExists = true;
            myDialog.Multiselect = true;
            if (myDialog.ShowDialog() == true)
            {
                TextBox_EmployeeFilePath.Text = myDialog.FileName;
                EmployeeFilePath = myDialog.FileName;
            }
        }

        private void Start_Click(object sender, RoutedEventArgs e)
        {
            if (String.IsNullOrWhiteSpace(TextBox_TeleoptiReportPath.Text) || String.IsNullOrWhiteSpace(TextBox_EmployeeFilePath.Text))
            {
                throw new ArgumentNullException("Выберите путь к файлу отчета.");
            }
            
            Main.Start();
        }

        private void TextBox_FirstDay_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            FirstDay = Convert.ToInt32(TextBox_FirstDay.Text);
        }

        private void TextBox_LastDay_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            LastDay = Convert.ToInt32(TextBox_LastDay.Text);
        }

        
    }
}
