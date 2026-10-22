// Title: Generate a shape inventory worksheet with type and cell coordinates using Aspose.Cells for .NET
// AI Prompts: Write C# code that uses Aspose.Cells to scan every worksheet, collect each shape's name, type, and upper‑left cell address, and write these details to a new worksheet named "ShapeReport". | Extend the program to also capture each shape's lower‑right cell address and add a corresponding column in the ShapeReport sheet. | Adapt the shape inventory solution to export the ShapeReport data as a CSV file instead of saving it within the original Excel workbook.
// Common Searches: Aspose.Cells C# list all drawing objects in an Excel workbook with their cell locations | how to retrieve shape type and upper left cell address using Aspose.Cells .NET | create a summary sheet of shapes for each worksheet with Aspose.Cells | export shape details from an Excel file to a new worksheet using Aspose.Cells | get shape coordinates in Excel via Aspose.Cells C# code
// Tags: Aspose.Cells enumerate worksheet shapes | extract shape type Aspose.Cells | shape coordinates Excel Aspose.Cells | generate shape report worksheet .NET | export shape data to CSV Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The program loads an existing Excel workbook, adds a new worksheet called "ShapeReport", iterates through all other worksheets, records each shape's name, type, and upper‑left cell address, writes this information into the report sheet, and saves the workbook to a new file.
class Program
{
    static void Main()
    {
        // Define input and output file paths
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Add a new worksheet to hold the shape report
            int reportSheetIndex = workbook.Worksheets.Add();
            var reportSheet = workbook.Worksheets[reportSheetIndex];
            reportSheet.Name = "ShapeReport";

            // Write header row
            reportSheet.Cells[0, 0].PutValue("Worksheet");
            reportSheet.Cells[0, 1].PutValue("Shape Name");
            reportSheet.Cells[0, 2].PutValue("Shape Type");
            reportSheet.Cells[0, 3].PutValue("Cell Address");

            int reportRow = 1;

            // Iterate through all worksheets except the report sheet itself
            foreach (Worksheet ws in workbook.Worksheets)
            {
                if (ws == reportSheet) continue;

                // Iterate through all shapes in the current worksheet
                foreach (Shape shape in ws.Shapes)
                {
                    // Get shape type as a readable string
                    string shapeType = shape.Type.ToString();

                    // Determine the cell address of the shape's upper‑left corner
                    string cellAddress = CellsHelper.CellIndexToName(shape.UpperLeftRow, shape.UpperLeftColumn);

                    // Populate the report row
                    reportSheet.Cells[reportRow, 0].PutValue(ws.Name);
                    reportSheet.Cells[reportRow, 1].PutValue(shape.Name);
                    reportSheet.Cells[reportRow, 2].PutValue(shapeType);
                    reportSheet.Cells[reportRow, 3].PutValue(cellAddress);
                    reportRow++;
                }
            }

            // Save the workbook with the new report sheet
            workbook.Save(outputPath);
            Console.WriteLine($"Shape report saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
