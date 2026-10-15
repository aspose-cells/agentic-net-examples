// Title: Import a 2D string array into an Excel worksheet and format the header row with a bold blue font and gray background using Aspose.Cells for .NET
// AI Prompts: Import a two‑dimensional string array into a worksheet at cell A1 and apply a style that makes the first row bold, blue, size 12 with a light‑gray fill using Aspose.Cells. | Create a reusable Style object that sets font color, weight, size, and background, then apply it only to the header range of the imported data in a .NET workbook. | Adjust the example to use an italic red font and add a thin black border around the header cells while keeping the data import unchanged.
// Common Searches: Aspose.Cells C# import two dimensional array into Excel | How to style header row after importing data with Aspose.Cells .NET | Set bold blue font and background color for first row in Aspose.Cells workbook | Apply StyleFlag for font and cell shading to a range in Aspose.Cells C#
// Tags: import 2d string array Aspose.Cells .NET | header row style Aspose.Cells | bold blue font Excel range C# | styleflag font cell shading Aspose | background fill header Aspose.Cells | reusable style object Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Demonstrates importing a two‑dimensional string array into a worksheet, creating a custom style with bold blue 12‑pt font and light‑gray fill, and applying it to the header row using Aspose.Cells for .NET.
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

            // Two‑dimensional array of strings (including header row)
            string[,] data = new string[,]
            {
                { "Header1", "Header2", "Header3" },
                { "Row1Col1", "Row1Col2", "Row1Col3" },
                { "Row2Col1", "Row2Col2", "Row2Col3" }
            };

            // Import the array into the worksheet starting at cell A1 (row 0, column 0)
            sheet.Cells.ImportTwoDimensionArray(data, 0, 0);

            // Create a custom style for the header row
            Style headerStyle = workbook.CreateStyle();
            Font headerFont = headerStyle.Font;
            headerFont.IsBold = true;          // Bold text
            headerFont.Size = 12;              // Font size 12
            headerFont.Color = Color.Blue;    // Font color blue

            // Optional: set background fill for the header cells
            headerStyle.ForegroundColor = Color.LightGray;
            headerStyle.Pattern = BackgroundType.Solid;

            // Define the range that covers the header row (first row, all columns)
            int columnCount = data.GetLength(1);
            AsposeRange headerRange = sheet.Cells.CreateRange(0, 0, 1, columnCount);

            // Apply the style to the header range (only font and shading)
            StyleFlag flag = new StyleFlag { Font = true, CellShading = true };
            headerRange.ApplyStyle(headerStyle, flag);

            // Save the workbook to a file
            workbook.Save("Output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
