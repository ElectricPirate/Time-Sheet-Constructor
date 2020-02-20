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
    public class ViewModel : INotifyPropertyChanged, IDataErrorInfo
    {
        private string employeeFilePath;
        public string EmployeeFilePath
        {
            get => employeeFilePath; 
            set 
            { 
                employeeFilePath = value;
                if (String.IsNullOrWhiteSpace(employeeFilePath))
                {
                    errors["EmployeeFilePath"] = "Адрес не может быть пустым";
                }
                else
                {
                    errors["EmployeeFilePath"] = null;
                }
            }
        }
        
        private string teleoptiReportPath;
        public string TeleoptiReportPath
        {
            get { return teleoptiReportPath; }
            set 
            {
                teleoptiReportPath = value;
                if (String.IsNullOrWhiteSpace(teleoptiReportPath))
                {
                    errors["TeleoptiReportPath"] = "Адрес не может быть пустым";
                }
                else
                {
                    errors["TeleoptiReportPath"] = null;
                }
            }
        }

        private string tableLayoutPath;
        public string TableLayoutPath
        {
            get => tableLayoutPath;
            set 
            { 
                tableLayoutPath = value;                
                if (String.IsNullOrWhiteSpace(tableLayoutPath))
                {
                    errors["TableLayoutPath"] = "Адрес не может быть пустым";
                }
                else
                {
                    errors["TableLayoutPath"] = null;
                }
            }
        }

        private int firstDay;
        public int FirstDay
        {
            get => firstDay;            
            set
            {
                firstDay = value;
                OnPropertyChanged("FirstDay");
                Main.FirstDay = firstDay;
                if (value <= 0 || value > 31 || value.ToString()==null)
                {
                    errors["FirstDay"] = "Некорректный день";
                }
                else
                {
                    errors["FirstDay"] = null;
                }
            }
        }

        private int lastDay;
        public int LastDay
        {
            get => lastDay;
            set
            {
                lastDay = value;
                OnPropertyChanged("LastDay");
                Main.LastDay = lastDay;
                if (value > 0 && value <= 31)
                {                    
                    errors["LastDay"] = null;
                }
                else
                {
                    errors["LastDay"] = "Некорректный день";
                }
            }
        }

        Dictionary<string, string> errors;

        public bool IsValid => !errors.Values.Any(x => x != null);

        IDialogService dialogService;

        public ViewModel(IDialogService dialogService)
        {
            this.dialogService = dialogService;
            errors = new Dictionary<string, string>();
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
                      teleoptiReportPath = dialogService.FilePath;
                      OnPropertyChanged("TeleoptiReportPath");
                      Main.TeleoptiReportPath = teleoptiReportPath;
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
                      employeeFilePath = dialogService.FilePath;
                      OnPropertyChanged("EmployeeFilePath");
                      Main.EmployeeFilePath = employeeFilePath;
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
                      tableLayoutPath = dialogService.FilePath;
                      OnPropertyChanged("TableLayoutPath");
                      Main.TableLayoutPath = tableLayoutPath;
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
                      if (IsValid)
                      {
                          Main.Start();
                      }
                  }));
            }
        }        

        public event PropertyChangedEventHandler PropertyChanged;
        public void OnPropertyChanged([CallerMemberName]string prop = "")
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(prop));
        }

        public string Error => throw new NotImplementedException();

        public string this[string columnName] => errors.ContainsKey(columnName) ? errors[columnName] : null;
    }
}
