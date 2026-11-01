// Title: Classify each worksheet in an Excel file as data‑only, shape‑only, mixed, or empty using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that iterates all worksheets, checks for any non‑empty cell values and any drawing objects, and prints a classification label for each sheet. | Create a reusable function that receives a Worksheet object and returns an enum (DataOnly, ShapeOnly, Mixed, Empty) based on cell content and shape count. | Enhance the classifier to also output the count of populated cells and the count of shapes found on each worksheet.
// Common Searches: how to determine if an Excel worksheet contains only data using Aspose.Cells C# | C# Aspose.Cells detect charts or pictures in a worksheet | classify Excel sheets as data only, shape only, or mixed with Aspose.Cells .NET | Aspose.Cells get used range and count shapes per worksheet | identify empty worksheets versus those with content in Aspose.Cells
// Tags: worksheet data detection Aspose.Cells | shape enumeration Excel Aspose.Cells | content classification worksheets .NET | used range analysis Aspose.Cells | mixed content detection Excel

using System;
using System.IO;
using Aspose.Cells;

namespace WorksheetContentClassifier
{
    // Loads an Excel workbook, scans each worksheet for any non‑empty cell values and any shapes, classifies the sheet as Data‑Only, Shape‑Only, Mixed, or Empty, and prints the result.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the input workbook
            string filePath = "input.xlsx";

            // Verify that the file exists to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Iterate through each worksheet in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    bool hasData = false;
                    bool hasShape = false;

                    // Determine the used range of the worksheet (use fully qualified type to avoid ambiguity)
                    Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;
                    int startRow = usedRange.FirstRow;
                    int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
                    int startCol = usedRange.FirstColumn;
                    int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                    // Scan cells for any non‑empty value
                    for (int row = startRow; row <= endRow && !hasData; row++)
                    {
                        for (int col = startCol; col <= endCol && !hasData; col++)
                        {
                            Cell cell = sheet.Cells[row, col];
                            if (cell.Value != null && !string.IsNullOrWhiteSpace(cell.Value.ToString()))
                            {
                                hasData = true;
                            }
                        }
                    }

                    // Check for any drawing objects (pictures, charts, etc.)
                    if (sheet.Shapes.Count > 0)
                    {
                        hasShape = true;
                    }

                    // Classify the worksheet based on data and shapes presence
                    string classification = hasData && !hasShape ? "Data-Only"
                                         : !hasData && hasShape ? "Shape-Only"
                                         : hasData && hasShape ? "Mixed"
                                         : "Empty";

                    // Output the classification result
                    Console.WriteLine($"Worksheet \"{sheet.Name}\": {classification}");
                }
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors during processing
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
