using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Runtime.InteropServices;
using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using System.Drawing.Imaging;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Table;

namespace Time_Sheet_Constructor.Model
{
    public class FileConstructor
    {
        public static string OutFilePath =
            @"C:\Users\vadim.turetskiy\Documents\Табель\Time sheet constructor\Табель.xlsx";

        public static void Create(ExcelPackage file)
        {
            var worksheet = file.Workbook.Worksheets.Add("Чистовик");
            CreateHead(worksheet);

            file.SaveAs(new FileInfo(OutFilePath));
        }

        private static void CreateHead(ExcelWorksheet sheet)
        {
            sheet.Column(1).SetTrueColumnWidth(3.25);
            sheet.Column(2).SetTrueColumnWidth(26.75);
            sheet.Column(3).SetTrueColumnWidth(5.5);

            for (int colIndex = 4; colIndex <= 35; colIndex++)
            {
                sheet.Column(colIndex).SetTrueColumnWidth(2.5);
            }

            sheet.Column(19).Width = 5.13;
            sheet.Column(36).Width = 5.13;
            sheet.Column(37).Width = 4.63;
            sheet.Column(38).Width = 5.13;
            sheet.Column(39).Width = 4.75;
            sheet.Column(40).Width = 5;
            sheet.Column(41).Width = 8;
            sheet.Column(42).Width = 3.75;
            sheet.Column(43).Width = 3.63;
            sheet.Column(44).Width = 3.63;
            sheet.Column(45).Width = 3.63;
            sheet.Column(46).Width = 2.13;
            sheet.Column(47).Width = 4.25;
            sheet.Column(48).Width = 3.63;
            sheet.Column(49).Width = 3.38;
            sheet.Column(50).Width = 3.25;
            sheet.Column(51).Width = 2.25;
            sheet.Column(52).Width = 2.63;
            sheet.Column(53).Width = 3.75;
            sheet.Column(54).Width = 3.25;
            sheet.Column(55).Width = 3.5;





        }

        private void SetWidht(int rowIndex)
        {
            var mdw = xlWorksheet.Workbook.MaxFontWidth;
            pixelHeight = (int)(worksheet.Row(rowIndex).Height / 0.75);
            pixelWidth = (int)decimal.Truncate(((256 * (decimal)worksheet.Column(columnIndex).Width +
                                                 decimal.Truncate(128 / (decimal)mdw)) / 256) * mdw);
        }
    }
}
