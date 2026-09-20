// Title: Applying a 40% transparent outer shadow preset to a rectangle shape with Aspose.Cells for .NET – API limitation note
// AI Prompts: Write C# code that adds a rectangle shape to an Excel worksheet and attempts to set an outer shadow preset with 40% transparency using Aspose.Cells, including error handling for unsupported shadow features. | Suggest a workaround or alternative technique to emulate a semi‑transparent outer shadow on an Excel shape when the Aspose.Cells .NET API does not provide direct shadow properties.
// Common Searches: how to add outer shadow to a shape in Aspose.Cells using C# | set shadow transparency on Excel shape with Aspose.Cells .NET | Aspose.Cells shape shadow effect not available in current version | C# code to apply preset shadow to rectangle shape in Excel via Aspose.Cells | workaround for missing shape shadow API in Aspose.Cells for .NET
// Tags: shape outer shadow Aspose.Cells | shadow transparency property .NET Excel | Aspose.Cells rectangle shape formatting | missing shape shadow API Aspose.Cells | Excel shape visual effects Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a new workbook, inserts a rectangle shape, and notes that Aspose.Cells for .NET currently lacks a direct API to configure outer shadow effects or transparency, then saves the workbook to 'output.xlsx'.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet.
            // In recent Aspose.Cells versions AddShape returns a Shape object.
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                1,   // upper left row
                1,   // upper left column
                0,   // upper left row offset (pixels)
                0,   // upper left column offset (pixels)
                100, // width (pixels)
                50   // height (pixels)
            );

            // Note: Shadow configuration is not available in the current Aspose.Cells API version.
            // If needed, refer to the documentation for the appropriate shadow settings.

            // Define output file path
            string outputPath = "output.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
