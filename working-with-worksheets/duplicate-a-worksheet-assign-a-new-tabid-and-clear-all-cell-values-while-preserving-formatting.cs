// Title: Duplicate an Excel worksheet, assign a new TabId, and clear all cell values while preserving formatting with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# duplicate worksheet and set custom TabId | clear data in copied worksheet but keep styles using Aspose.Cells | how to assign a new TabId to a duplicated sheet in Aspose.Cells | remove all cell values from a copied sheet while preserving formatting Aspose.Cells | use ClearContents on used range after copying a worksheet with Aspose.Cells
// Tags: duplicate worksheet with new TabId Aspose.Cells | clear worksheet contents preserve formatting Aspose.Cells | copy sheet and reset data Aspose.Cells C# | assign custom TabId to worksheet Aspose.Cells | ClearContents on used range Aspose.Cells

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;

namespace WorksheetDuplicateAndClear
{
    // The example loads an existing workbook (or creates one), duplicates the first worksheet, renames it, assigns a new TabId by adding an offset, clears all cell values in the copied sheet using ClearContents on the used range while retaining original formatting, and saves the modified workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Input and output file paths (can be overridden via command‑line arguments)
            string inputPath = args.Length > 0 ? args[0] : "input.xlsx";
            string outputPath = args.Length > 1 ? args[1] : "output.xlsx";

            try
            {
                Workbook workbook;

                // Load existing workbook if it exists; otherwise create a sample workbook.
                if (File.Exists(inputPath))
                {
                    Console.WriteLine($"Loading workbook from '{inputPath}'.");
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    Console.WriteLine($"Input file '{inputPath}' not found. Creating a sample workbook.");
                    workbook = new Workbook();

                    // Prepare a simple source worksheet with some data and formatting.
                    Worksheet sourceSample = workbook.Worksheets[0];
                    sourceSample.Name = "Source";

                    // Header with bold font and yellow background.
                    sourceSample.Cells["A1"].PutValue("Product");
                    sourceSample.Cells["B1"].PutValue("Price");
                    Style headerStyle = sourceSample.Cells["A1"].GetStyle();
                    headerStyle.Font.IsBold = true;
                    headerStyle.ForegroundColor = Color.Yellow;
                    headerStyle.Pattern = BackgroundType.Solid;
                    sourceSample.Cells["A1"].SetStyle(headerStyle);
                    sourceSample.Cells["B1"].SetStyle(headerStyle);

                    // Sample data rows.
                    sourceSample.Cells["A2"].PutValue("Apple");
                    sourceSample.Cells["B2"].PutValue(1.25);
                    sourceSample.Cells["A3"].PutValue("Orange");
                    sourceSample.Cells["B3"].PutValue(0.95);
                }

                // Ensure there is at least one worksheet to copy.
                if (workbook.Worksheets.Count == 0)
                {
                    throw new InvalidOperationException("The workbook does not contain any worksheets to duplicate.");
                }

                // Reference to the original worksheet (first worksheet in this example).
                Worksheet originalWorksheet = workbook.Worksheets[0];
                Console.WriteLine($"Original worksheet name: '{originalWorksheet.Name}', TabId: {originalWorksheet.TabId}");

                // Duplicate the worksheet. AddCopy returns the index of the newly created sheet.
                int copiedIndex = workbook.Worksheets.AddCopy(0);
                Worksheet copiedWorksheet = workbook.Worksheets[copiedIndex];
                copiedWorksheet.Name = originalWorksheet.Name + "_Copy";

                // Assign a new TabId to the duplicated worksheet.
                // Here we simply add an offset to guarantee uniqueness.
                int newTabId = originalWorksheet.TabId + 1000;
                copiedWorksheet.TabId = newTabId;
                Console.WriteLine($"Copied worksheet name: '{copiedWorksheet.Name}', new TabId: {copiedWorksheet.TabId}");

                // Clear all cell values while preserving formatting.
                // Cells.ClearContents requires a range; we clear the used range.
                int maxRow = copiedWorksheet.Cells.MaxDataRow;
                int maxColumn = copiedWorksheet.Cells.MaxDataColumn;
                if (maxRow >= 0 && maxColumn >= 0)
                {
                    // Clear from first cell (0,0) covering all used rows/columns.
                    copiedWorksheet.Cells.ClearContents(0, 0, maxRow + 1, maxColumn + 1);
                }
                Console.WriteLine("All cell values in the copied worksheet have been cleared, formatting retained.");

                // Save the modified workbook.
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("An error occurred during processing:");
                Console.Error.WriteLine(ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
