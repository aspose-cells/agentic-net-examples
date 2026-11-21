// Title: C# routine to extract shape hyperlinks and generate a worksheet report of source cells and URLs using Aspose.Cells
// Common Searches: how to list all shape hyperlinks with their anchor cells in an Excel file using Aspose.Cells for .NET | Aspose.Cells C# extract hyperlink address from shapes and write to a new worksheet | generate a hyperlink report for shapes in a workbook with Aspose.Cells | retrieve top‑left cell reference of a shape that contains a hyperlink in C# | save Excel workbook after adding a shape hyperlink summary sheet using Aspose.Cells
// Tags: shape hyperlink extraction Aspose.Cells | shape anchor cell retrieval Aspose.Cells | hyperlink summary worksheet generation C# | enumerate workbook shapes Aspose.Cells | save workbook with additional report Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example loads an existing Excel workbook, adds a new worksheet named "HyperlinkReport", writes header columns, then iterates through every worksheet and each shape on those sheets. For shapes that contain a non‑empty Hyperlink, it captures the shape's top‑left cell reference (e.g., "A1") and the hyperlink address, records the worksheet name, source cell, and URL in the report sheet, and finally saves the workbook with the generated report.
class ShapeHyperlinkExtractor
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output_with_hyperlink_report.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Add a new worksheet for the hyperlink report
            int reportIndex = workbook.Worksheets.Add();
            Worksheet reportSheet = workbook.Worksheets[reportIndex];
            reportSheet.Name = "HyperlinkReport";

            // Write header row
            reportSheet.Cells[0, 0].PutValue("Worksheet");
            reportSheet.Cells[0, 1].PutValue("SourceCell");
            reportSheet.Cells[0, 2].PutValue("TargetURL");

            int reportRow = 1;

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through each shape on the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    try
                    {
                        // Check if the shape contains a hyperlink with a valid address
                        if (shape.Hyperlink != null && !string.IsNullOrEmpty(shape.Hyperlink.Address))
                        {
                            // Determine the top‑left cell that anchors the shape
                            int row = shape.UpperLeftRow;
                            int column = shape.UpperLeftColumn;
                            string sourceCell = sheet.Cells[row, column].Name; // e.g., "A1"

                            // Record the information in the report sheet
                            reportSheet.Cells[reportRow, 0].PutValue(sheet.Name);
                            reportSheet.Cells[reportRow, 1].PutValue(sourceCell);
                            reportSheet.Cells[reportRow, 2].PutValue(shape.Hyperlink.Address);
                            reportRow++;
                        }
                    }
                    catch (Exception exShape)
                    {
                        // Log shape‑specific errors but continue processing other shapes
                        Console.WriteLine($"Error processing shape on sheet \"{sheet.Name}\": {exShape.Message}");
                    }
                }
            }

            // Save the workbook with the added report
            workbook.Save(outputPath);
            Console.WriteLine($"Hyperlink report saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
