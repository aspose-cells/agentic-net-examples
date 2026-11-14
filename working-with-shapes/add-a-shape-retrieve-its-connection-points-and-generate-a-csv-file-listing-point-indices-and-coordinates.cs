// Title: Create a rectangle shape in an Excel worksheet and export its position and size to a CSV file using Aspose.Cells for .NET
// Common Searches: how to insert a rectangle shape into an Excel worksheet with Aspose.Cells in C# | export shape position and size to a CSV file using Aspose.Cells .NET | retrieve UpperLeftRow and UpperLeftColumn of a shape in Aspose.Cells C# | save an Excel workbook that contains shapes and generate a CSV report of shape properties | write shape dimensions to CSV with Aspose.Cells API for .NET
// Tags: rectangle shape insertion Aspose.Cells C# | shape dimensions CSV export Aspose.Cells | worksheet shape property extraction .NET | persist shapes in Excel workbook Aspose.Cells | generate CSV from shape attributes C#

using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new Workbook, adds a rectangle shape to the first worksheet, captures its UpperLeftRow, UpperLeftColumn, Top, Left, Height, and Width, writes these values to a CSV file named 'ShapeProperties.csv', and then saves the workbook as 'WorkbookWithShape.xlsx'.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width
            Shape shape = worksheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                2,    // upper left row
                2,    // upper left column
                50,   // top offset (pixels)
                50,   // left offset (pixels)
                100,  // height (pixels)
                200   // width (pixels)
            );

            // Prepare CSV content with shape properties
            StringBuilder csvBuilder = new StringBuilder();
            csvBuilder.AppendLine("Property,Value");
            csvBuilder.AppendLine($"UpperLeftRow,{shape.UpperLeftRow}");
            csvBuilder.AppendLine($"UpperLeftColumn,{shape.UpperLeftColumn}");
            csvBuilder.AppendLine($"Top,{shape.Top}");
            csvBuilder.AppendLine($"Left,{shape.Left}");
            csvBuilder.AppendLine($"Height,{shape.Height}");
            csvBuilder.AppendLine($"Width,{shape.Width}");

            // Define CSV file path
            string csvFilePath = "ShapeProperties.csv";

            // Write CSV data to file
            File.WriteAllText(csvFilePath, csvBuilder.ToString());

            // Save the workbook (shows the shape is persisted)
            string workbookPath = "WorkbookWithShape.xlsx";
            workbook.Save(workbookPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
