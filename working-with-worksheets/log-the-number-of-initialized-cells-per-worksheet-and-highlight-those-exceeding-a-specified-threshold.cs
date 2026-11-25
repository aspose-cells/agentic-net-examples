// Title: Count initialized cells per worksheet and highlight cells that exceed a threshold using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# count non‑empty cells per worksheet | highlight cells when initialized cell count exceeds a limit using Aspose.Cells | log worksheet cell usage statistics with Aspose.Cells .NET | iterate used range with MaxDataRow and MaxDataColumn in Aspose.Cells | create yellow style for cells that surpass a count threshold in C#
// Tags: initialized cell count Aspose.Cells | threshold based cell highlighting .NET | MaxDataRow MaxDataColumn iteration Aspose.Cells | apply style to cells exceeding count Aspose.Cells | worksheet cell usage logging C#

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace CellInitializerLogger
{
    // The example loads an Excel workbook, creates a yellow highlight style, iterates each worksheet using MaxDataRow and MaxDataColumn to count non‑null cells, logs the count per sheet, applies the highlight style to cells once the count passes a defined threshold, and saves the modified workbook.
    class Program
    {
        // Threshold for number of initialized cells. Worksheets exceeding this will be highlighted.
        private const int CellCountThreshold = 1000;

        // Input and output file paths.
        private const string InputFilePath = "input.xlsx";
        private const string OutputFilePath = "output.xlsx";

        static void Main(string[] args)
        {
            try
            {
                // Ensure the input file exists; create an empty workbook if it does not.
                if (!File.Exists(InputFilePath))
                {
                    Console.WriteLine($"Input file \"{InputFilePath}\" not found. Creating a new empty workbook.");
                    var emptyWorkbook = new Workbook();
                    emptyWorkbook.Save(InputFilePath, SaveFormat.Xlsx);
                }

                // Load the workbook from the input file.
                Workbook workbook = new Workbook(InputFilePath);
                Console.WriteLine($"Workbook loaded successfully from \"{InputFilePath}\".");

                // Prepare a style for highlighting cells that exceed the threshold.
                Style highlightStyle = workbook.CreateStyle();
                highlightStyle.ForegroundColor = Color.Yellow;
                highlightStyle.Pattern = BackgroundType.Solid;

                // Process each worksheet.
                foreach (Worksheet worksheet in workbook.Worksheets)
                {
                    Cells cells = worksheet.Cells;

                    // Determine the maximum data row and column.
                    int maxRow = cells.MaxDataRow;       // Returns -1 if no data.
                    int maxCol = cells.MaxDataColumn;    // Returns -1 if no data.

                    // If there is no data, report and continue.
                    if (maxRow < 0 || maxCol < 0)
                    {
                        Console.WriteLine($"Worksheet \"{worksheet.Name}\": No initialized cells found.");
                        continue;
                    }

                    // Count initialized (non‑empty) cells.
                    int initializedCellCount = 0;
                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxCol; col++)
                        {
                            Cell cell = cells[row, col];

                            // A cell is considered initialized if its type is not Null.
                            if (cell.Type != CellValueType.IsNull)
                            {
                                initializedCellCount++;

                                // Apply highlight immediately once the threshold is exceeded.
                                if (initializedCellCount > CellCountThreshold)
                                {
                                    cell.SetStyle(highlightStyle);
                                }
                            }
                        }
                    }

                    // Log the result.
                    Console.WriteLine($"Worksheet \"{worksheet.Name}\": Initialized cells = {initializedCellCount}.");

                    // Optional: highlight the whole used range if the threshold is exceeded.
                    if (initializedCellCount > CellCountThreshold)
                    {
                        Console.WriteLine($"  -> Highlighted cells exceeding threshold of {CellCountThreshold}.");
                    }
                }

                // Save the modified workbook.
                workbook.Save(OutputFilePath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved with highlights to \"{OutputFilePath}\".");
            }
            catch (Exception ex)
            {
                // Report any unexpected errors.
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
