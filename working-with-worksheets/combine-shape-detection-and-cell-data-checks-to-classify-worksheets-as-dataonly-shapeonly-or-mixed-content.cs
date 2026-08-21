// Title: Classify Excel Worksheets as Data‑Only, Shape‑Only, Mixed or Empty with Aspose.Cells for .NET (C#)
// Description: Loads a workbook, iterates each worksheet, detects shapes via the Shapes collection, scans cells up to MaxDataRow/MaxDataColumn for non‑empty values, and assigns a content type (Data‑Only, Shape‑Only, Mixed, Empty). Results are printed to the console and the workbook can be saved.
// Keywords: Aspose.Cells worksheet classification | detect shapes Aspose.Cells | check cell data Aspose.Cells | C# Excel shape detection | Excel worksheet content type | MaxDataRow Aspose.Cells | MaxDataColumn Aspose.Cells | Aspose.Cells .NET | Excel sheet empty detection
// Common Searches: Aspose.Cells how to find worksheets with only charts | C# detect if Excel sheet contains data using Aspose.Cells | classify Excel worksheets by content Aspose.Cells | identify empty worksheets in a workbook with Aspose.Cells | determine mixed content worksheets Aspose.Cells
// Developer Intent: Identify whether each worksheet contains shapes, data, both, or nothing and label it accordingly.
// Use Cases: Audit large workbooks and generate a summary that lists each sheet as Data‑Only, Shape‑Only, Mixed or Empty. | Skip shape‑only worksheets when extracting tabular data for migration or reporting scripts. | Apply custom export or formatting rules based on the sheet's content type (e.g., export data‑only sheets to CSV). | Create automated documentation of workbook structure for compliance or quality checks.
// AI Prompts: Create a reusable method that returns an enum (DataOnly, ShapeOnly, Mixed, Empty) for a Worksheet using Aspose.Cells. | Rewrite the classification logic with LINQ and parallel processing to improve performance. | Add detailed logging that records the number of shapes, data cells, and the final classification for each worksheet. | Extend the example to write the classification results into a new summary worksheet within the same workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace WorksheetClassificationDemo
{
    // Loads a workbook, iterates each worksheet, detects shapes via the Shapes collection, scans cells up to MaxDataRow/MaxDataColumn for non‑empty values, and assigns a content type (Data‑Only, Shape‑Only, Mixed, Empty). Results are printed to the console and the workbook can be saved.
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your file path)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Determine if the worksheet contains any shapes
                bool hasShapes = sheet.Shapes.Count > 0;

                // Determine if the worksheet contains any data (non‑empty cells)
                bool hasData = false;

                // Use the maximum used row and column indices to limit the scan
                int maxRow = sheet.Cells.MaxDataRow;      // Last row with data
                int maxCol = sheet.Cells.MaxDataColumn;   // Last column with data

                for (int row = 0; row <= maxRow && !hasData; row++)
                {
                    for (int col = 0; col <= maxCol && !hasData; col++)
                    {
                        Cell cell = sheet.Cells[row, col];
                        if (cell != null && cell.Type != CellValueType.IsNull)
                        {
                            hasData = true;
                        }
                    }
                }

                // Classify the worksheet based on the presence of shapes and data
                string classification;
                if (hasData && !hasShapes)
                {
                    classification = "Data‑Only";
                }
                else if (!hasData && hasShapes)
                {
                    classification = "Shape‑Only";
                }
                else if (hasData && hasShapes)
                {
                    classification = "Mixed Content";
                }
                else
                {
                    classification = "Empty";
                }

                Console.WriteLine($"Worksheet \"{sheet.Name}\": {classification}");
            }

            // Optionally save the workbook (e.g., after modifications)
            workbook.Save("output.xlsx");
        }
    }
}
