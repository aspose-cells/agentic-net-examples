using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Enable strict Open XML compliance (ISO/IEC 29500:2008 Strict)
        workbook.Settings.Compliance = OoxmlCompliance.Iso29500_2008_Strict;

        // (Optional) Add a sample value to demonstrate the workbook is usable
        workbook.Worksheets[0].Cells["A1"].PutValue("Strict compliance enabled");

        // Save the workbook using a default file name
        workbook.Save("StrictComplianceWorkbook.xlsx");
    }
}