// Title: Export every worksheet from a merged Excel workbook to individual CSV files using Aspose.Cells for .NET
// Common Searches: C# code to save each sheet of a merged workbook as separate CSV with Aspose.Cells | How to loop through worksheets and export to CSV using Aspose.Cells .NET | Extract worksheets to CSV files from a combined Excel file in C# | Aspose.Cells export multiple sheets to CSV programmatically | Create one‑sheet workbook per worksheet and save as CSV in .NET
// Tags: export worksheets to CSV Aspose.Cells | iterate workbook sheets C# | single-sheet workbook creation Aspose.Cells | save worksheet as CSV .NET | merged Excel workbook processing Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads a merged Excel workbook, iterates through all its worksheets, creates a new one‑sheet workbook for each sheet, copies the original content while preserving the sheet name, and saves each workbook as a CSV file named after the worksheet, handling missing files and export errors.
class ExportWorksheetsToCsv
{
    static void Main()
    {
        try
        {
            // Path to the merged workbook file (XLSX, XLS, etc.)
            string workbookPath = "MergedWorkbook.xlsx";

            // Verify that the source workbook exists
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: Workbook file \"{workbookPath}\" not found.");
                return;
            }

            // Load the workbook using Aspose.Cells
            Workbook workbook = new Workbook(workbookPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    // Create a new workbook that will contain only the current worksheet
                    Workbook singleSheetWorkbook = new Workbook();

                    // The new workbook contains one default sheet; copy the source sheet into it
                    Worksheet destSheet = singleSheetWorkbook.Worksheets[0];
                    destSheet.Name = sheet.Name;          // Preserve original sheet name
                    destSheet.Copy(sheet);                // Copy contents and formatting

                    // Define CSV file name based on worksheet name
                    string csvFileName = $"{sheet.Name}.csv";

                    // Save the single‑sheet workbook as CSV
                    singleSheetWorkbook.Save(csvFileName, SaveFormat.Csv);

                    Console.WriteLine($"Exported \"{sheet.Name}\" to \"{csvFileName}\".");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to export sheet \"{sheet.Name}\": {ex.Message}");
                }
            }

            Console.WriteLine("Export completed.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
