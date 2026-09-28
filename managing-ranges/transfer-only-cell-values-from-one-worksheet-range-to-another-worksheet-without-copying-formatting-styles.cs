// Title: Copy only cell values from a source range to a destination range on another worksheet with Aspose.Cells for .NET (C#)
// AI Prompts: Read the values from A1:E10 on Sheet1 and write them to the range starting at C6 on Sheet2 using Cell.PutValue in C#. | Open source.xlsx, ensure destination.xlsx contains Sheet2, then iterate through the source CellArea and transfer only the cell values to the target worksheet while preserving its existing formatting. | Use Aspose.Cells to copy data between two workbooks by looping over rows and columns and calling PutValue for each cell to achieve a values‑only paste.
// Common Searches: asp.net aspose.cells copy only values between worksheets c# | c# aspose.cells transfer range values without formatting | how to paste values only from one workbook to another using Aspose.Cells | aspose.cells putvalue example for copying data between sheets
// Tags: value-only range copy Aspose.Cells C# | transfer cell values between worksheets Aspose.Cells | copy values without formatting Aspose.Cells | Cell.PutValue method range transfer C# | worksheet value transfer Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsValueTransfer
{
    // The program loads a source workbook, creates or opens a destination workbook, ensures Sheet2 exists, and then copies only the cell values from the source range A1:E10 on Sheet1 to the destination range starting at C6 on Sheet2 using Cell.PutValue, preserving any existing formatting in the destination before saving the updated file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Verify source file exists
                const string sourcePath = "source.xlsx";
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file \"{sourcePath}\" not found.");
                    return;
                }

                // Load the source workbook
                Workbook sourceWorkbook = new Workbook(sourcePath);

                // Verify destination file exists; if not, create a new workbook
                const string destinationPath = "destination.xlsx";
                Workbook destinationWorkbook;
                if (File.Exists(destinationPath))
                {
                    destinationWorkbook = new Workbook(destinationPath);
                }
                else
                {
                    destinationWorkbook = new Workbook(); // creates a default workbook with one sheet
                }

                // Get the source worksheet (fallback to first sheet if name not found)
                Worksheet sourceSheet = sourceWorkbook.Worksheets["Sheet1"] ??
                                        sourceWorkbook.Worksheets[0];

                // Get the destination worksheet; add if missing
                Worksheet destinationSheet = destinationWorkbook.Worksheets["Sheet2"];
                if (destinationSheet == null)
                {
                    int index = destinationWorkbook.Worksheets.Add();
                    destinationSheet = destinationWorkbook.Worksheets[index];
                    destinationSheet.Name = "Sheet2";
                }

                // Define the source range (A1:E10)
                CellArea sourceRange = new CellArea
                {
                    StartRow = 0,      // Row 1 (0‑based)
                    StartColumn = 0,   // Column A (0‑based)
                    EndRow = 9,        // Row 10
                    EndColumn = 4      // Column E
                };

                // Destination top‑left cell (C6)
                int destStartRow = 5;      // Row 6 (0‑based)
                int destStartColumn = 2;   // Column C (0‑based)

                // Transfer only values from source range to destination range
                for (int i = 0; i <= sourceRange.EndRow - sourceRange.StartRow; i++)
                {
                    for (int j = 0; j <= sourceRange.EndColumn - sourceRange.StartColumn; j++)
                    {
                        Cell srcCell = sourceSheet.Cells[sourceRange.StartRow + i, sourceRange.StartColumn + j];
                        Cell destCell = destinationSheet.Cells[destStartRow + i, destStartColumn + j];
                        destCell.PutValue(srcCell.Value);
                    }
                }

                // Save the updated destination workbook
                const string outputPath = "destination_updated.xlsx";
                destinationWorkbook.Save(outputPath);
                Console.WriteLine($"Values transferred successfully. Saved as \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
