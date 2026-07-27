// Title: C# – Load a Partially Corrupted Excel Workbook with Warning Callback, RepairLoad & DataExtractLoad (Aspose.Cells)
// Description: Demonstrates how to open a potentially damaged .xlsx file in .NET, capture each load warning via a custom IWarningCallback, enable RepairLoad and DataExtractLoad to recover usable data, skip unnecessary shapes, and read the first few cells from the recovered worksheet.
// Keywords: Aspose.Cells | C# | .NET | load corrupted workbook | warning callback | RepairLoad | DataExtractLoad | IgnoreUselessShapes | CheckDataValid false | recoverable Excel data | GitHub example | openxlsx damaged file | Excel file repair
// Common Searches: Aspose.Cells load corrupted xlsx with warnings | C# warning callback for Excel workbook load | Enable RepairLoad in Aspose.Cells | How to extract data from a damaged Excel file using Aspose | Skip shapes when opening a corrupted workbook Aspose.Cells
// Developer Intent: Open a damaged Excel file, log all load warnings, activate repair and data‑extract modes, and continue processing the recovered content.
// Use Cases: Log warning types and messages to diagnose file integrity issues. | Recover cell values from a partially corrupted sheet for further analysis. | Improve load performance on broken files by ignoring non‑essential shapes.
// AI Prompts: Write C# code that opens a corrupted .xlsx with Aspose.Cells, uses a custom IWarningCallback to print warnings, enables RepairLoad and DataExtractLoad, and iterates over the first 10 rows of column B. | Explain how WarningCallback, IgnoreUselessShapes, and CheckDataValid affect the loading of a damaged workbook in Aspose.Cells. | Generate a GitHub‑style README snippet describing this example and its required NuGet packages.

using System;
using System.IO;
using Aspose.Cells;

namespace Example
{
    // Demonstrates how to open a potentially damaged .xlsx file in .NET, capture each load warning via a custom IWarningCallback, enable RepairLoad and DataExtractLoad to recover usable data, skip unnecessary shapes, and read the first few cells from the recovered worksheet.
    public class CustomWarningCallback : IWarningCallback
    {
        // Called for each warning generated during load
        public void Warning(WarningInfo warningInfo)
        {
            Console.WriteLine($"Warning Type: {warningInfo.Type}, Description: {warningInfo.Description}");
        }
    }

    public class LoadCorruptedWorkbookDemo
    {
        public static void Run()
        {
            string filePath = "corrupted.xlsx";

            // Prevent FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            try
            {
                // Set up warning callback
                IWarningCallback warningCallback = new CustomWarningCallback();

                // Configure load options
                LoadOptions loadOptions = new LoadOptions
                {
                    WarningCallback = warningCallback,
                    IgnoreUselessShapes = true,
                    CheckDataValid = false
                };

                // Load the workbook with the specified options
                Workbook workbook = new Workbook(filePath, loadOptions);

                // Enable repair and data‑extract modes
                workbook.Settings.RepairLoad = true;
                workbook.Settings.DataExtractLoad = true;

                // Example processing: read first few cells from the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Console.WriteLine("First 5 cells in column A (recoverable content):");
                for (int row = 0; row < 5; row++)
                {
                    Console.WriteLine($"A{row + 1}: {sheet.Cells[row, 0].StringValue}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading workbook: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            LoadCorruptedWorkbookDemo.Run();
        }
    }
}
