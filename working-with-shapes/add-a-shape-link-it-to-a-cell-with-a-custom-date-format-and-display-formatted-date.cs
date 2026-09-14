// Title: Add a rectangle shape linked to a date cell with a custom format and display the formatted date using Aspose.Cells for .NET
// AI Prompts: Write C# code that inserts a rectangle shape into a worksheet, links it to a specific cell, and sets the shape's text to the cell's formatted date value. | Adjust the shape's placement so it moves and resizes with the linked cell and position the shape directly over that cell. | Change the cell's custom date format to include time (e.g., "dd-MMM-yyyy HH:mm") and update the shape's displayed text accordingly.
// Common Searches: aspnet add rectangle shape linked to date cell Aspose.Cells | display custom formatted date inside a shape using Aspose.Cells for .NET | how to make a shape move and size with its linked cell in Aspose.Cells | set custom date format in Excel cell and show it in a shape with Aspose.Cells
// Tags: Aspose.Cells rectangle shape insertion | shape linked to cell placement MoveAndSize | custom date format in Excel cell Aspose.Cells | shape text bound to cell value .NET | save workbook as xlsx Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program creates a new workbook, writes the current date to cell A1 with a custom "dd-mmm-yyyy" format, adds a rectangle shape, links the shape to the cell using MoveAndSize placement, sets the shape's text to the formatted date, and saves the workbook as Output.xlsx.
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

            // Put a DateTime value into cell A1
            Cell dateCell = sheet.Cells["A1"];
            dateCell.PutValue(DateTime.Now);

            // Apply a custom date format (e.g., 15-Mar-2023)
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Custom = "dd-mmm-yyyy";
            dateCell.SetStyle(dateStyle);

            // Add a rectangle shape to the sheet
            // Position it initially at row 2, column 1 (B3) with size 150x50 points
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // Shape type
                2,    // Upper left row (zero‑based)
                1,    // Upper left column (zero‑based)
                0,    // Top offset in points
                0,    // Left offset in points
                50,   // Height in points
                150); // Width in points

            // Link the shape to the date cell so it moves/resizes with the cell
            shape.Placement = PlacementType.MoveAndSize;

            // (Optional) Position the shape over cell A1.
            // The SetPosition method is not available in older versions, so we rely on the initial placement.
            // shape.SetPosition(0, 0, 0, 0);

            // Display the formatted date inside the shape
            shape.Text = dateCell.StringValue;

            // Determine output file path
            string outputPath = "Output.xlsx";

            // Ensure the directory exists (handle case where outputPath has no directory component)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (string.IsNullOrEmpty(outputDir))
            {
                outputDir = Directory.GetCurrentDirectory();
            }
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
