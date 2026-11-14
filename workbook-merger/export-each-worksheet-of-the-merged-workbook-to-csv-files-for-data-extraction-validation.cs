// Title: Export each worksheet from a merged Excel workbook to separate CSV files using Aspose.Cells for .NET
// Common Searches: C# code to save every sheet of a merged .xlsx as individual CSV files with Aspose.Cells | How to iterate worksheets in Aspose.Cells and export them to CSV in .NET | Aspose.Cells .NET export multiple worksheets to separate CSV files with UTF-8 encoding | Create CSV files for each worksheet after merging workbooks using Aspose.Cells | Save each worksheet as CSV using TxtSaveOptions in Aspose.Cells C#
// Tags: Aspose.Cells TxtSaveOptions CSV export | C# iterate workbook worksheets | save single-sheet workbook as CSV | UTF-8 encoding for CSV with Aspose.Cells | output folder for CSV exports in .NET | export merged workbook sheets to separate CSV files

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace ExportWorksheetsToCsv
{
    // The example loads a merged Excel file, ensures an output directory exists, loops through all worksheets, copies each sheet into a new single‑sheet workbook, builds a safe CSV filename, configures TxtSaveOptions with UTF‑8 encoding, and saves each sheet as an individual CSV file while logging successes and errors.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the merged workbook (adjust as needed)
            string workbookPath = "merged.xlsx";

            // Verify that the source workbook exists
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: The workbook file '{workbookPath}' was not found.");
                return;
            }

            // Directory where CSV files will be saved
            string outputDirectory = "CsvExports";

            // Ensure the output directory exists
            if (!Directory.Exists(outputDirectory))
            {
                Directory.CreateDirectory(outputDirectory);
            }

            Workbook workbook = null;
            try
            {
                // Load the merged workbook
                workbook = new Workbook(workbookPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    // Create a new workbook containing only the current worksheet
                    Workbook singleSheetWorkbook = new Workbook();

                    // Get the default (first) worksheet in the new workbook
                    Worksheet newSheet = singleSheetWorkbook.Worksheets[0];

                    // Copy the source worksheet's content into the new workbook
                    newSheet.Copy(sheet);
                    newSheet.Name = sheet.Name;

                    // Build a safe CSV file name (replace invalid file chars)
                    string safeSheetName = string.Join("_", sheet.Name.Split(Path.GetInvalidFileNameChars()));
                    string csvFilePath = Path.Combine(outputDirectory, $"{safeSheetName}.csv");

                    // Configure CSV save options
                    TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv)
                    {
                        // Example: set delimiter and encoding if needed
                        // Separator = ',',
                        Encoding = Encoding.UTF8
                    };

                    // Save the single-sheet workbook as CSV
                    singleSheetWorkbook.Save(csvFilePath, saveOptions);

                    Console.WriteLine($"Exported worksheet '{sheet.Name}' to CSV: {csvFilePath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to export worksheet '{sheet.Name}': {ex.Message}");
                }
            }

            Console.WriteLine("All worksheets have been processed.");
        }
    }
}
