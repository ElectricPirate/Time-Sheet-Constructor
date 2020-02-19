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


        //private void TeleoptiReportPath_Click(object sender, RoutedEventArgs e)
        //{
        //    OpenFileDialog myDialog = new OpenFileDialog();
        //    myDialog.Filter = "Книга Excel(*.xlsx;*.xls)|*.xlsx;*.xls" + "|Все файлы (*.*)|*.* ";
        //    myDialog.CheckFileExists = true;
        //    myDialog.Multiselect = true;
        //    if (myDialog.ShowDialog() == true)
        //    {
        //        TextBox_TeleoptiReportPath.Text = myDialog.FileName;
        //        TeleoptiReportPath = myDialog.FileName;
        //    }
        //}

        //private void EmployeeFilePath_Click(object sender, RoutedEventArgs e)
        //{
        //    OpenFileDialog myDialog = new OpenFileDialog();
        //    myDialog.Filter = "Книга Excel(*.xlsx;*.xls)|*.xlsx;*.xls" + "|Все файлы (*.*)|*.* ";
        //    myDialog.CheckFileExists = true;
        //    myDialog.Multiselect = true;
        //    if (myDialog.ShowDialog() == true)
        //    {
        //        TextBox_EmployeeFilePath.Text = myDialog.FileName;
        //        EmployeeFilePath = myDialog.FileName;
        //    }
        //}

        //private void TableLayoutPath_Click(object sender, RoutedEventArgs e)
        //{
        //    OpenFileDialog myDialog = new OpenFileDialog();
        //    myDialog.Filter = "Книга Excel(*.xlsx;*.xls)|*.xlsx;*.xls" + "|Все файлы (*.*)|*.* ";
        //    myDialog.CheckFileExists = true;
        //    myDialog.Multiselect = true;
        //    if (myDialog.ShowDialog() == true)
        //    {
        //        TextBox_TableLayoutPath.Text = myDialog.FileName;
        //        TableLayoutPath = myDialog.FileName;
        //    }
        //}

        //private void Start_Click(object sender, RoutedEventArgs e)
        //{
        //    if (String.IsNullOrWhiteSpace(TextBox_TeleoptiReportPath.Text) || String.IsNullOrWhiteSpace(TextBox_EmployeeFilePath.Text) ||String.IsNullOrWhiteSpace(TextBox_TableLayoutPath.Text))
        //    {
        //        MessageBox.Show("Выберите путь к файлу отчета.", "Ошибка");
        //    }
        //    else if (String.IsNullOrEmpty(TextBox_FirstDay.Text) || String.IsNullOrEmpty(TextBox_LastDay.Text))
        //    {
        //        MessageBox.Show("Введите корректные даты.", "Ошибка");
        //    }
        //    else
        //    {
        //        Main.Start();
        //    }
        //}

        //private void TextBox_FirstDay_OnTextChanged(object sender, TextChangedEventArgs e)
        //{
        //    var result = 0;

        //    if(!int.TryParse(TextBox_FirstDay.Text, out result))
        //    {
        //        TextBox_FirstDay.Text = "";
        //    }
        //    else if (Convert.ToInt32(TextBox_FirstDay.Text) < 1 || Convert.ToInt32(TextBox_FirstDay.Text) > 31)
        //    {
        //        TextBox_FirstDay.Text = "";
        //    }
        //    else
        //    {
        //        FirstDay = Convert.ToInt32(TextBox_FirstDay.Text);
        //    }
        //}

        //private void TextBox_LastDay_OnTextChanged(object sender, TextChangedEventArgs e)
        //{
        //    var result = 0;

        //    if (!int.TryParse(TextBox_LastDay.Text, out result))
        //    {
        //        TextBox_LastDay.Text = "";
        //    }
        //    else if (Convert.ToInt32(TextBox_LastDay.Text) < 1 || Convert.ToInt32(TextBox_LastDay.Text) > 31)
        //    {
        //        TextBox_LastDay.Text = "";
        //    }
        //    else
        //    {
        //        LastDay = Convert.ToInt32(TextBox_LastDay.Text);
        //    }
        //}



    }
}
