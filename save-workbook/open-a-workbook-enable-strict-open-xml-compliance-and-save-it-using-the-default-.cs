using System;
using Aspose.Cells;

namespace AsposeCellsStrictComplianceDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (uses Workbook() constructor rule)
            Workbook workbook = new Workbook();

            // Access workbook settings and enable strict OOXML compliance
            // (uses Workbook.Settings and WorkbookSettings.Compliance property rule)
            workbook.Settings.Compliance = OoxmlCompliance.Iso29500_2008_Strict;

            // Optionally add some data to demonstrate the workbook is functional
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Strict OOXML Compliance Enabled");

            // Save the workbook using a default file name (uses Workbook.Save(string) rule)
            workbook.Save("StrictComplianceWorkbook.xlsx");
        }
    }
}