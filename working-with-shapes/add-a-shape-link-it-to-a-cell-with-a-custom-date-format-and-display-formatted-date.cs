// Title: Create a rectangle shape linked to a cell with a custom date format in Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add rectangle shape and bind it to a cell formula | how to display a formatted date from a worksheet cell inside a shape using Aspose.Cells | set custom date format for a cell and link shape text to the cell in Aspose.Cells .NET | link shape text to cell B2 with custom date formatting Aspose.Cells C#
// Tags: rectangle shape insertion Aspose.Cells | bind shape text to cell formula Aspose.Cells | apply custom date format to cell Aspose.Cells | shape font color and size Aspose.Cells | save workbook containing shapes Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, writes the current date to cell B2, applies the custom "dd-mmm-yyyy" format, inserts a rectangle shape, sets the shape's text to the formula "=B2" so it shows the formatted date, customizes the shape's font color and size, and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Put a date value into cell B2
            Cell dateCell = worksheet.Cells["B2"];
            dateCell.PutValue(DateTime.Now);

            // Apply a custom date format to the cell
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Custom = "dd-mmm-yyyy";
            dateCell.SetStyle(dateStyle);

            // Add a rectangle shape to the worksheet
            Shape shape = null;
            try
            {
                // AddShape returns a Shape object in recent Aspose.Cells versions
                shape = worksheet.Shapes.AddShape(
                    MsoDrawingType.Rectangle, // shape type
                    5,   // upper left row
                    0,   // upper left column
                    5,   // top offset (pixels)
                    0,   // left offset (pixels)
                    200, // height (pixels)
                    50   // width (pixels)
                );
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to add shape: {ex.Message}");
            }

            if (shape != null)
            {
                // Link the shape's text to the cell B2 so it displays the formatted date
                shape.Text = "=B2";

                // Optional: format the shape's text appearance
                shape.Font.Color = Color.Blue;
                shape.Font.Size = 12;
            }

            // Define output file path
            string outputPath = "output.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (string.IsNullOrEmpty(outputDir))
            {
                outputDir = Directory.GetCurrentDirectory();
            }
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
