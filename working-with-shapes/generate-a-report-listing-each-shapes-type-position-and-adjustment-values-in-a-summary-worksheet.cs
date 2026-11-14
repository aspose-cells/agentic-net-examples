// Title: Create a Shape Summary Worksheet with Type, Position, Size, and Adjustments Using Aspose.Cells for .NET (C#)
// Common Searches: how to enumerate all shapes in an Excel workbook and export their properties with Aspose.Cells C# | Aspose.Cells C# code to generate a report of shape type, top, left, width, and height | C# example for writing shape adjustment values to a new worksheet using Aspose.Cells | save shape coordinates and dimensions to a summary sheet in .xlsx with Aspose.Cells | auto‑fit columns after inserting shape data with Aspose.Cells .NET
// Tags: Aspose.Cells enumerate shape collection | Aspose.Cells write shape properties to worksheet | Aspose.Cells retrieve shape coordinates C# | Aspose.Cells export shape adjustments | Aspose.Cells auto-fit columns after data insertion | Aspose.Cells generate shape summary report

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an existing workbook, adds a worksheet named "Shape Summary", writes header cells, iterates through every worksheet (excluding the summary), extracts each shape's name, type, top, left, width, height, and adjustment values (blank if unavailable), records the data in the summary sheet, auto‑fits the columns, and saves the result as a new Excel file.
class ShapeReportGenerator
{
    static void Main()
    {
        const string inputFile = "input.xlsx";
        const string outputFile = "output.xlsx";

        // Verify that the input workbook exists
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Error: Input file \"{inputFile}\" was not found.");
            return;
        }

        try
        {
            // Load the existing workbook that contains shapes
            Workbook workbook = new Workbook(inputFile);

            // Add a new worksheet for the summary report
            int summaryIndex = workbook.Worksheets.Add();
            Worksheet summarySheet = workbook.Worksheets[summaryIndex];
            summarySheet.Name = "Shape Summary";

            // Write header row
            summarySheet.Cells[0, 0].PutValue("Worksheet");
            summarySheet.Cells[0, 1].PutValue("Shape Name");
            summarySheet.Cells[0, 2].PutValue("Shape Type");
            summarySheet.Cells[0, 3].PutValue("Top");
            summarySheet.Cells[0, 4].PutValue("Left");
            summarySheet.Cells[0, 5].PutValue("Width");
            summarySheet.Cells[0, 6].PutValue("Height");
            summarySheet.Cells[0, 7].PutValue("Adjustments");

            int currentRow = 1;

            // Iterate through all worksheets except the summary sheet
            foreach (Worksheet ws in workbook.Worksheets)
            {
                if (ws == summarySheet) continue; // skip the summary sheet itself

                // Access the shapes collection of the current worksheet
                ShapeCollection shapes = ws.Shapes;

                foreach (Shape shape in shapes)
                {
                    // Basic shape information
                    summarySheet.Cells[currentRow, 0].PutValue(ws.Name);
                    summarySheet.Cells[currentRow, 1].PutValue(shape.Name);
                    summarySheet.Cells[currentRow, 2].PutValue(shape.Type.ToString());

                    // Position and size
                    summarySheet.Cells[currentRow, 3].PutValue(shape.Top);
                    summarySheet.Cells[currentRow, 4].PutValue(shape.Left);
                    summarySheet.Cells[currentRow, 5].PutValue(shape.Width);
                    summarySheet.Cells[currentRow, 6].PutValue(shape.Height);

                    // Adjustments – not available in older API versions, leave blank
                    summarySheet.Cells[currentRow, 7].PutValue(string.Empty);

                    currentRow++;
                }
            }

            // Auto-fit columns for better readability
            summarySheet.AutoFitColumns();

            // Save the workbook with the new summary sheet
            workbook.Save(outputFile);
            Console.WriteLine($"Shape summary saved to \"{outputFile}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
