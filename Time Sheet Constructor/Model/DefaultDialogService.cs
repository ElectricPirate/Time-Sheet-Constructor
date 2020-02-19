using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace Time_Sheet_Constructor.Model
{
    public class DefaultDialogService : IDialogService
    {
        public string FilePath { get; set; }

        public bool OpenFileDialog()
        {
            var openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Книга Excel(*.xlsx;*.xls)|*.xlsx;*.xls" + "|Все файлы (*.*)|*.* ";
            openFileDialog.CheckFileExists = true;

            if (openFileDialog.ShowDialog() == true)
            {
                FilePath = openFileDialog.FileName;
                return true;
            }

            return false;
        }

        public bool SaveFileDialog()
        {
            var saveFileDialog = new SaveFileDialog();

            if (saveFileDialog.ShowDialog() == true)
            {
                FilePath = saveFileDialog.FileName;
                return true;
            }

            return false;
        }

        public void ShowMessage(string message)
        {
            MessageBox.Show(message);
        }
    }
}
