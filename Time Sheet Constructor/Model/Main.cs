using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;


namespace Time_Sheet_Constructor.Model
{
    public static class Main
    {
        public static string TeleoptiReportPath { get; set; }
        public static string EmployeeFilePath { get; set; }
        public static string TableLayoutPath { get; set; }
        public static int FirstDay { get; set; }
        public static int LastDay { get; set; }

        public static void Start()
        {     
            var existingFile = new FileInfo(TeleoptiReportPath);
            
            var table = new ExcelPackage(existingFile);

            var Persons = new List<Person>();
            var data = new FileParser(table);
            Persons = data.GetData();               
            var ParseIDs = new EmpoyeeIDParser(Persons, EmployeeFilePath);
            var PeronsWithIDs = ParseIDs.Parse();
            var DraftData = new ExportDraft(TableLayoutPath, PeronsWithIDs, FirstDay, LastDay);

            DraftData.Write();

        }
    }
}
