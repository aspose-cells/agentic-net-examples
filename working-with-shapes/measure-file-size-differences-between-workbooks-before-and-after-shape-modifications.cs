// Title: Measure XLSX file size change after adding a rectangle shape with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# get workbook size before inserting a shape | compare Excel file size after adding rectangle shape using Aspose.Cells | how to calculate XLSX size increase when a shape is added in .NET | determine memory stream length of workbook with Aspose.Cells after modifications | measure file size difference of Excel workbook before and after shape insertion
// Tags: measure workbook size Aspose.Cells | shape insertion file size impact XLSX | calculate Excel file growth C# | rectangle shape addition Aspose.Cells | memory stream size comparison Aspose.Cells | XLSX byte count change after shape

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, records its size by saving to a MemoryStream, adds a rectangle shape with text, saves the modified workbook to another MemoryStream, and then outputs the original size, modified size, and the byte difference to illustrate how shape insertion affects XLSX file size.
class ShapeSizeComparison
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Save the original workbook to a memory stream to get its size
        int originalSize;
        using (MemoryStream originalStream = new MemoryStream())
        {
            workbook.Save(originalStream, SaveFormat.Xlsx);
            originalSize = (int)originalStream.Length;
        }

        // ---- Shape modifications start ----
        // Add a rectangle shape to the worksheet
        Shape rect = sheet.Shapes.AddShape(
            MsoDrawingType.Rectangle, // shape type
            1,   // upper left row
            1,   // upper left column
            0,   // top offset in pixels
            0,   // left offset in pixels
            100, // width in pixels
            50   // height in pixels
        );
        rect.Text = "Sample Shape";
        // ---- Shape modifications end ----

        // Save the modified workbook to another memory stream to get its new size
        int modifiedSize;
        using (MemoryStream modifiedStream = new MemoryStream())
        {
            workbook.Save(modifiedStream, SaveFormat.Xlsx);
            modifiedSize = (int)modifiedStream.Length;
        }

        // Calculate and display the size difference
        int sizeDifference = modifiedSize - originalSize;
        Console.WriteLine($"Original size: {originalSize} bytes");
        Console.WriteLine($"Modified size: {modifiedSize} bytes");
        Console.WriteLine($"Size difference: {sizeDifference} bytes");
    }
}
