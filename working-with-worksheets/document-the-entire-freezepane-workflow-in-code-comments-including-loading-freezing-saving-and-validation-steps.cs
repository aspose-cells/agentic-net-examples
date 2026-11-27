// Title: C# Aspose.Cells tutorial: load or generate workbook, apply FreezePanes at C3, save as Xlsx, and verify frozen panes
// Common Searches: C# Aspose.Cells code to freeze rows and columns at cell C3 | How to automatically create a sample Excel file when the input file is missing using Aspose.Cells | Saving an Excel workbook with frozen panes and then reopening it for validation in .NET | Applying FreezePanes method with row and column counts in Aspose.Cells for C# | Best practices for error handling in Aspose.Cells freeze pane operations
// Tags: apply FreezePanes at specific cell Aspose.Cells C# | generate sample Excel workbook Aspose.Cells .NET | save workbook with frozen panes Xlsx Aspose.Cells | reopen and verify frozen pane configuration Aspose.Cells | exception handling for Excel freeze pane workflow C#

using System;
using System.IO;
using Aspose.Cells;

namespace FreezePaneDemo
{
    // The example demonstrates a complete freeze‑pane workflow: it loads an existing workbook or creates a sample one when the input file is absent, accesses the first worksheet, applies FreezePanes at cell C3 freezing the first two rows and columns, saves the modified workbook as an Xlsx file, and finally reopens the saved file to confirm the frozen pane settings, with comprehensive error handling throughout.
    class Program
    {
        static void Main(string[] args)
        {
            // -----------------------------------------------------------------
            // 1. Define input and output file paths.
            // -----------------------------------------------------------------
            string inputPath = "input.xlsx";
            string outputPath = "output_frozen.xlsx";

            try
            {
                // -----------------------------------------------------------------
                // 2. Load the workbook.
                //    If the input file does not exist, create a simple workbook
                //    with sample data so the demo can run without hidden failures.
                // -----------------------------------------------------------------
                Workbook workbook;
                if (File.Exists(inputPath))
                {
                    Console.WriteLine($"Loading workbook from \"{inputPath}\"...");
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a sample workbook.");
                    workbook = new Workbook(); // creates a workbook with one default worksheet
                    Worksheet sampleSheet = workbook.Worksheets[0];
                    sampleSheet.Name = "SampleData";

                    // Populate some sample data
                    sampleSheet.Cells["A1"].PutValue("Item");
                    sampleSheet.Cells["B1"].PutValue("Quantity");
                    sampleSheet.Cells["A2"].PutValue("Apple");
                    sampleSheet.Cells["B2"].PutValue(10);
                    sampleSheet.Cells["A3"].PutValue("Banana");
                    sampleSheet.Cells["B3"].PutValue(20);
                    sampleSheet.Cells["A4"].PutValue("Total");
                    sampleSheet.Cells["B4"].Formula = "=SUM(B2:B3)";

                    // Save the sample workbook so the user can see the original file
                    workbook.Save(inputPath, SaveFormat.Xlsx);
                    Console.WriteLine($"Sample workbook saved as \"{inputPath}\".");
                }

                // -----------------------------------------------------------------
                // 3. Access the worksheet on which we want to freeze panes.
                //    For this demo we use the first worksheet.
                // -----------------------------------------------------------------
                Worksheet worksheet = workbook.Worksheets[0];
                Console.WriteLine($"Working on worksheet \"{worksheet.Name}\" (index {worksheet.Index}).");

                // -----------------------------------------------------------------
                // 4. Apply freeze panes.
                //    FreezePanes(string cellName, int totalRows, int totalColumns)
                //    - cellName: the top‑left cell that will remain scrollable (e.g., "C3").
                //    - totalRows: number of rows to freeze above the cell.
                //    - totalColumns: number of columns to freeze to the left of the cell.
                // -----------------------------------------------------------------
                string freezeCell = "C3";
                int rowsToFreeze = 2;      // Freeze rows 1‑2 (zero‑based index 0‑1)
                int columnsToFreeze = 2;   // Freeze columns A‑B (zero‑based index 0‑1)

                Console.WriteLine($"Applying FreezePanes at cell \"{freezeCell}\" with {rowsToFreeze} rows and {columnsToFreeze} columns frozen.");
                worksheet.FreezePanes(freezeCell, rowsToFreeze, columnsToFreeze);

                // -----------------------------------------------------------------
                // 5. Save the workbook with the freeze‑pane settings.
                // -----------------------------------------------------------------
                Console.WriteLine($"Saving modified workbook to \"{outputPath}\"...");
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine("Save completed.");

                // -----------------------------------------------------------------
                // 6. Validation step (optional):
                //    Re‑open the saved workbook to ensure it loads without error.
                // -----------------------------------------------------------------
                Console.WriteLine("Validating saved workbook...");
                Workbook validationWb = new Workbook(outputPath);
                Worksheet validationWs = validationWb.Worksheets[0];
                Console.WriteLine($"Validation succeeded: Worksheet \"{validationWs.Name}\" loaded successfully with freeze‑pane settings.");
            }
            catch (Exception ex)
            {
                // -----------------------------------------------------------------
                // 7. Error handling: report any unexpected issues.
                // -----------------------------------------------------------------
                Console.Error.WriteLine("An error occurred during the freeze‑pane workflow:");
                Console.Error.WriteLine(ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
                Environment.Exit(1);
            }
        }
    }
}
