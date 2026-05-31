using System;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsExportWorksheetToTxt
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and access the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Name = "DataSheet";

            // Populate the worksheet with sample data
            worksheet.Cells["A1"].PutValue("Name");
            worksheet.Cells["B1"].PutValue("Age");
            worksheet.Cells["A2"].PutValue("John");
            worksheet.Cells["B2"].PutValue(30);
            worksheet.Cells["A3"].PutValue("Jane");
            worksheet.Cells["B3"].PutValue(25);

            // If there are multiple worksheets and you want to export a specific one,
            // set it as the active sheet. Here we keep the first sheet active.
            workbook.Worksheets.ActiveSheetIndex = 0;

            // Configure TxtSaveOptions for tab-delimited UTF-8 output
            TxtSaveOptions saveOptions = new TxtSaveOptions
            {
                Separator = '\t',          // Tab delimiter
                Encoding = Encoding.UTF8, // UTF-8 encoding
                ExportAllSheets = false   // Export only the active sheet
            };

            // Export the active worksheet to a TXT file
            string outputPath = "ExportedWorksheet.txt";
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Worksheet exported successfully to '{outputPath}'.");
        }
    }
}