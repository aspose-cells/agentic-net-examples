using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;   // Alias to avoid conflict with System.Range

namespace AsposeCellsRangeDuplicate
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (no template file needed)
                Workbook workbook = new Workbook();

                // Get the first worksheet (source) and name it
                Worksheet sourceSheet = workbook.Worksheets[0];
                sourceSheet.Name = "Source";

                // Populate source range with values, a formula in the last column and a simple style
                Cells srcCells = sourceSheet.Cells;
                for (int row = 0; row < 5; row++)
                {
                    for (int col = 0; col < 5; col++)
                    {
                        // Put a numeric value
                        srcCells[row, col].PutValue(row * 10 + col);

                        // Add a SUM formula in column E (index 4) of each row
                        if (col == 4)
                        {
                            srcCells[row, col].Formula = $"=SUM(A{row + 1}:D{row + 1})";
                        }

                        // Apply a basic style
                        Style style = workbook.CreateStyle();
                        style.Font.Name = "Arial";
                        style.Font.Size = 12;
                        style.Font.IsBold = (col % 2 == 0);
                        srcCells[row, col].SetStyle(style);
                    }
                }

                // Add a new worksheet that will receive the duplicated range
                Worksheet destSheet = workbook.Worksheets[workbook.Worksheets.Add()];
                destSheet.Name = "Destination";

                // Define source and destination ranges (both 5x5)
                AsposeRange sourceRange = srcCells.CreateRange(0, 0, 5, 5);
                AsposeRange destRange = destSheet.Cells.CreateRange(0, 0, 5, 5);

                // Set paste options to copy everything (data, formulas, formats, etc.)
                PasteOptions pasteOptions = new PasteOptions
                {
                    PasteType = PasteType.All
                };

                // Perform the copy with the specified options
                destRange.Copy(sourceRange, pasteOptions);

                // Save the workbook
                string outputPath = "RangeDuplicateDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}