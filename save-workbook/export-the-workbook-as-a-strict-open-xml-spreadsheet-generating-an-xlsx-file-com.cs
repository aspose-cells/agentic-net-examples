using System;
using Aspose.Cells;

namespace StrictOoxmlExportDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (uses the provided creation rule)
            Workbook workbook = new Workbook();

            // Set the OOXML compliance level to ISO/IEC 29500:2008 Strict
            workbook.Settings.Compliance = OoxmlCompliance.Iso29500_2008_Strict;

            // Add sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Strict OOXML Export");
            sheet.Cells["A2"].PutValue(DateTime.Now);

            // Save the workbook as an XLSX file (uses the provided save rule)
            workbook.Save("StrictCompliance.xlsx");
        }
    }
}