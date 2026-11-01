// Title: Add a TextBox shape with a clickable hyperlink to an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add textbox shape with hyperlink to Excel file | create clickable shape linking to external URL in .xlsx using Aspose.Cells | set hyperlink address on a shape in Aspose.Cells .NET example | format textbox shape as blue underlined link with Aspose.Cells | insert shape with custom text and URL into worksheet using Aspose.Cells C#
// Tags: textbox shape with hyperlink Aspose.Cells | shape hyperlink address property C# | blue underlined font styling for shape Aspose.Cells | insert clickable shape into .xlsx worksheet | custom text in shape Aspose.Cells example

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a new workbook, adds a TextBox shape at a specific cell, sets custom display text, applies blue underlined font styling to mimic a hyperlink, assigns an external URL via the shape's Hyperlink.Address property, and saves the result as an XLSX file using Aspose.Cells for .NET.
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

            // Add a TextBox shape at row 2, column 2 (zero‑based indices)
            // Parameters: upper left row, upper left column, upper left offset X, upper left offset Y, width, height
            int upperLeftRow = 1;          // Row index (0‑based)
            int upperLeftColumn = 1;       // Column index (0‑based)
            int upperLeftRowOffset = 0;
            int upperLeftColumnOffset = 0;
            int width = 200;               // Width in points
            int height = 50;               // Height in points

            Shape shape = sheet.Shapes.AddTextBox(
                upperLeftRow,
                upperLeftColumn,
                upperLeftRowOffset,
                upperLeftColumnOffset,
                width,
                height);

            // Set the custom text displayed inside the shape
            shape.Text = "Click here for additional resources";

            // Apply basic formatting to make it look like a hyperlink
            shape.Font.Color = Color.Blue;
            shape.Font.Underline = FontUnderlineType.Single;

            // Configure the hyperlink (setting the address is sufficient for a URL)
            shape.Hyperlink.Address = "https://www.example.com";

            // Save the workbook to a file
            workbook.Save("ShapeWithHyperlink.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
