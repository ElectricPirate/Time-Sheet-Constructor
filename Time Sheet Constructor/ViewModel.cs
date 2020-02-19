using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using Time_Sheet_Constructor.Model;

namespace Time_Sheet_Constructor
{
    public class ViewModel : INotifyPropertyChanged
    {
        private string _employeeFilePath;
        public string EmployeeFilePath
        {
            get { return _employeeFilePath; }
            set 
            { 
                _employeeFilePath = value;                
            }
        }
        
        private string _teleoptiReportPath;
        public string TeleoptiReportPath
        {
            get { return _teleoptiReportPath; }
            set 
            {
                _teleoptiReportPath = value;                
            }
        }

        private string _tableLayoutPath;
        public string TableLayoutPath
        {
            get { return _tableLayoutPath; }
            set 
            { 
                _tableLayoutPath = value;                
            }
        }

        public static int FirstDay { get; set; }

        public static int LastDay { get; set; }

        IDialogService dialogService;

        public ViewModel(IDialogService dialogService)
        {
            this.dialogService = dialogService;
        }

        private RelayCommand openTeleoptiReportPathCommand;
        public RelayCommand OpenTeleoptiReportPathCommand
        {
            get
            {
                return openTeleoptiReportPathCommand ??
                  (openTeleoptiReportPathCommand = new RelayCommand(obj =>
                  {                      
                      this.dialogService.OpenFileDialog();
                      _teleoptiReportPath = dialogService.FilePath;
                      OnPropertyChanged("TeleoptiReportPath");
                      Main.TeleoptiReportPath = _teleoptiReportPath;
                  }));
            }
        }

        private RelayCommand openEmployeeFilePathCommand;
        public RelayCommand OpenEmployeeFilePathCommand
        {
            get
            {
                return openEmployeeFilePathCommand ??
                  (openEmployeeFilePathCommand = new RelayCommand(obj =>
                  {
                      this.dialogService.OpenFileDialog();
                      _employeeFilePath = dialogService.FilePath;
                      OnPropertyChanged("EmployeeFilePath");
                      EmpoyeeIDParser.EmployeeFilePath_xls = _employeeFilePath;
                  }));
            }
        }

        private RelayCommand openTableLayoutPathCommand;
        public RelayCommand OpenTableLayoutPathCommand
        {
            get
            {
                return openTableLayoutPathCommand ??
                  (openTableLayoutPathCommand = new RelayCommand(obj =>
                  {
                      this.dialogService.OpenFileDialog();
                      _tableLayoutPath = dialogService.FilePath;
                      OnPropertyChanged("TableLayoutPath");
                      ExportDraft.TableLayoutPath = _tableLayoutPath;
                  }));
            }
        }

        private RelayCommand startCommand;
        public RelayCommand StartCommand
        {
            get
            {
                return startCommand ??
                  (startCommand = new RelayCommand(obj =>
                  {
                      Main.Start();
                  }));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        public void OnPropertyChanged([CallerMemberName]string prop = "")
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(prop));
        }        
    }
}
