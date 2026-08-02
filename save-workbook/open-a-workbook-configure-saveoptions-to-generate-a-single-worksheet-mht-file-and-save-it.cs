using System;
using Aspose.Cells;

namespace AsposeCellsMhtExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file to be opened
            string sourcePath = "input.xlsx";

            // Open the existing workbook
            Workbook workbook = new Workbook(sourcePath);

            // Create HtmlSaveOptions for MHTML format
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.MHtml);

            // Configure options to generate a single‑worksheet MHT file
            saveOptions.SaveAsSingleFile = true;               // Save as a single file
            saveOptions.ExportActiveWorksheetOnly = true;      // Export only the active worksheet

            // Optionally set the active sheet index if a specific sheet is required
            // workbook.Worksheets.ActiveSheetIndex = 0; // default is the first sheet

            // Save the workbook as MHTML using the configured options
            string outputPath = "output.mht";
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook saved as single‑worksheet MHT file: {outputPath}");
        }
    }
}