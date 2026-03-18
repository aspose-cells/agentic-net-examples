using Aspose.Cells;
using System;

class StrictOoxmlSaveDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set the OOXML compliance level to ISO/IEC 29500:2008 Strict
        workbook.Settings.Compliance = OoxmlCompliance.Iso29500_2008_Strict;

        // Add sample data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Strict OOXML Compliance Demo");

        // Create OOXML save options (default options are sufficient)
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();

        // Save the workbook as a strict OOXML XLSX file using the save options
        workbook.Save("StrictCompliance.xlsx", saveOptions);
    }
}