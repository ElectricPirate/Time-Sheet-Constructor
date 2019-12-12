using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using OfficeOpenXml;

namespace Time_Sheet_Constructor.Model
{
    public class Main
    {
        public static void Start()
        {
            var existingFile = new FileInfo(MainWindow.TeleoptiReportPath);
            var table = new ExcelPackage(existingFile);

            
            var Persons = FileParser.GetData(table);

            Persons = EmpoyeeIDParser.Parse(Persons);

            ExportDraft.Write(Persons);

        }
    }
}
