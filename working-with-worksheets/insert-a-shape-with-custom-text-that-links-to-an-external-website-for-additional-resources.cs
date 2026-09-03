// Title: Add a rectangle shape with custom text and an external hyperlink to an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Create a new workbook, insert a rectangle shape at a specific cell, set its displayed text, and attach a hyperlink to https://www.example.com using the Aspose.Cells C# API. | Generate an Excel file where a shape acts as a clickable link by configuring Shape.Text and Shape.Hyperlink.Address with Aspose.Cells in .NET. | Write C# code that adds a rectangle shape to the first worksheet, assigns a custom caption, links it to an external URL, and saves the workbook as an Xlsx file.
// Common Searches: Aspose.Cells C# add rectangle shape with hyperlink to Excel file | how to set hyperlink on a shape using Aspose.Cells .NET | create clickable shape in Excel with Aspose.Cells API | example of inserting a shape with custom text and link in C# Aspose.Cells
// Tags: shape insertion with hyperlink Aspose.Cells | custom text on Excel shape .NET | hyperlink property on Aspose.Cells shape | save workbook containing linked shape Xlsx | Aspose.Cells shape API usage example

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a new workbook, accesses the first worksheet, adds a rectangle shape at row 1, column 0 with defined dimensions, sets its text to "Click here for more resources", assigns a hyperlink pointing to https://www.example.com, and saves the file as ShapeWithLink.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape.
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                1, 0,                     // upper‑left row and offset
                1, 0,                     // upper‑left column and offset
                200, 50);                 // width and height

            // Set custom text inside the shape.
            shape.Text = "Click here for more resources";

            // Assign a hyperlink to the shape by setting its Address property.
            shape.Hyperlink.Address = "https://www.example.com";

            // Save the workbook.
            string outputPath = "ShapeWithLink.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
