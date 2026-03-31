using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set the OOXML compliance level to ISO/IEC 29500:2008 Strict
        WorkbookSettings settings = workbook.Settings;
        settings.Compliance = OoxmlCompliance.Iso29500_2008_Strict;

        // Add sample data (optional)
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells[0, 0].PutValue("Strict OOXML compliance");

        // Save the workbook in XLSX format with strict compliance
        workbook.Save("StrictCompliance.xlsx");
    }
}