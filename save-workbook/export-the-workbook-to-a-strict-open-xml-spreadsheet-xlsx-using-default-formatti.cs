using System;
using Aspose.Cells;

namespace AsposeCellsExportStrictXlsx
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (default formatting)
            Workbook workbook = new Workbook();

            // Optionally add some data to demonstrate the workbook content
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["B1"].PutValue("World");

            // Set the OOXML compliance level to ISO/IEC 29500:2008 Strict
            workbook.Settings.Compliance = OoxmlCompliance.Iso29500_2008_Strict;

            // Save the workbook as a strict Open XML Spreadsheet (.xlsx) using default settings
            workbook.Save("StrictOutput.xlsx", SaveFormat.Xlsx);

            // Clean up
            workbook.Dispose();

            Console.WriteLine("Workbook saved as strict Open XML Spreadsheet: StrictOutput.xlsx");
        }
    }
}