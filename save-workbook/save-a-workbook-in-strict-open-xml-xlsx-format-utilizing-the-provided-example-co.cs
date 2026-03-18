using System;
using Aspose.Cells;

namespace AsposeCellsStrictSaveDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Set the OOXML compliance level to ISO/IEC 29500:2008 Strict
            workbook.Settings.Compliance = OoxmlCompliance.Iso29500_2008_Strict;

            // Add some data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Strict OOXML Demo");
            sheet.Cells["B2"].PutValue(123);

            // Save the workbook in XLSX format (strict compliance)
            workbook.Save("StrictComplianceDemo.xlsx");
        }
    }
}