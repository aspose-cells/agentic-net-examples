// Title: Classify Excel Worksheets as Data‑Only, Shape‑Only, Mixed or Empty using Aspose.Cells for .NET (C#)
// Description: Loads a workbook, iterates each worksheet, checks the Shapes collection and scans the cell range defined by MaxDataRow/MaxDataColumn to detect any non‑null values, then labels the sheet as Data‑Only, Shape‑Only, Mixed or Empty and writes the result to the console. Includes optional LoadOptions.IgnoreUselessShapes for faster processing.
// Keywords: Aspose.Cells worksheet classification | C# detect shapes in Excel | check non‑empty cells Aspose.Cells | Excel sheet content type detection | .NET Excel mixed content | IgnoreUselessShapes option
// Common Searches: how to identify data‑only vs shape‑only worksheets with Aspose.Cells | C# code to classify Excel sheets as data, shape, mixed or empty | detect empty worksheets and shapes in a .NET workbook | Aspose.Cells MaxDataRow MaxDataColumn example
// Developer Intent: Determine whether each worksheet contains cell data, drawing shapes, both, or nothing.
// Use Cases: Create a summary report that lists every worksheet and its content type for workbook cleanup. | Skip shape‑only sheets when exporting data to CSV or a database, processing only data‑only sheets. | Route worksheets to different pipelines (e.g., data extraction vs. image extraction) based on their classification.
// AI Prompts: Generate a reusable method that returns "Data‑Only", "Shape‑Only", "Mixed" or "Empty" for a given Worksheet using Aspose.Cells. | Modify the sample to also count shapes and include the count in the console output. | Write unit tests that verify classification for worksheets with only data, only shapes, both, and none.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace WorksheetClassificationDemo
{
    // Loads a workbook, iterates each worksheet, checks the Shapes collection and scans the cell range defined by MaxDataRow/MaxDataColumn to detect any non‑null values, then labels the sheet as Data‑Only, Shape‑Only, Mixed or Empty and writes the result to the console. Includes optional LoadOptions.IgnoreUselessShapes for faster processing.
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your file path)
            string inputPath = "input.xlsx";
            LoadOptions loadOptions = new LoadOptions
            {
                // Optional: ignore duplicate/overlapping shapes for faster processing
                IgnoreUselessShapes = true
            };
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Iterate through each worksheet and classify its content
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                bool hasShapes = sheet.Shapes.Count > 0;
                bool hasData = false;

                // Determine if the sheet contains any non‑empty cells
                int maxRow = sheet.Cells.MaxDataRow;      // -1 if no data
                int maxCol = sheet.Cells.MaxDataColumn;   // -1 if no data

                if (maxRow >= 0 && maxCol >= 0)
                {
                    for (int r = 0; r <= maxRow && !hasData; r++)
                    {
                        for (int c = 0; c <= maxCol; c++)
                        {
                            Cell cell = sheet.Cells[r, c];
                            // Cell.Type == CellValueType.IsNull indicates an empty cell
                            if (cell != null && cell.Type != CellValueType.IsNull)
                            {
                                hasData = true;
                                break;
                            }
                        }
                    }
                }

                // Classify the worksheet based on presence of data and shapes
                string classification;
                if (hasData && hasShapes)
                    classification = "Mixed Content";
                else if (hasData)
                    classification = "Data‑Only";
                else if (hasShapes)
                    classification = "Shape‑Only";
                else
                    classification = "Empty";

                Console.WriteLine($"Worksheet \"{sheet.Name}\": {classification}");
            }

            // Optionally, save the workbook after processing (not required for classification)
            // workbook.Save("output.xlsx");
        }
    }
}
