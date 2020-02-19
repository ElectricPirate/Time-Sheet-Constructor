using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;


namespace Time_Sheet_Constructor.Model
{
    public static class Main
    {
        public static string TeleoptiReportPath { get; set; }             
    
        public static void Start()
        {     
            var existingFile = new FileInfo(TeleoptiReportPath);
            
            var table = new ExcelPackage(existingFile);

            var Persons = new List<Person>();
            var data = new FileParser(table);
            Persons = data.GetData();               
            var ParseIDs = new EmpoyeeIDParser(Persons);
            var PeronsWithIDs = ParseIDs.Parse(Persons);
            

            ExportDraft.Write(PeronsWithIDs);

        }
    }
}
