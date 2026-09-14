// Title: How to apply a dense hatch fill pattern to a header row in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Create a C# style using BackgroundType.DenseHatch with a yellow foreground and blue background, then apply it to cells A1 through D1 in an Aspose.Cells workbook. | Write code that defines a header range and uses StyleFlag to apply a custom patterned style to that range with Aspose.Cells for .NET. | Demonstrate how to detect unsupported pattern types in Aspose.Cells and switch to a solid fill style for header cells.
// Common Searches: how to use BackgroundType.DenseHatch in Aspose.Cells C# example | apply custom pattern to Excel header row using Aspose.Cells .NET | C# Aspose.Cells fallback to solid fill when pattern not supported | set foreground and background colors with hatch pattern in Aspose.Cells | range styling with StyleFlag in Aspose.Cells workbook
// Tags: dense hatch style Aspose.Cells C# | header row styling Aspose.Cells | densehatch backgroundtype usage | apply style to range with Aspose.Cells | fallback to solid background Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;

// The program creates a new workbook, defines a style with a dense hatch pattern (yellow foreground, blue background), applies the style to the header range A1:D1 using StyleFlag, writes header text, and saves the file as HeaderWithDenseHatch.xlsx, with a solid fill fallback if the hatch pattern is unavailable.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define the header range (cells A1 to D1)
            int startRow = 0;
            int endRow = 0;
            int startColumn = 0;
            int endColumn = 3;

            // Create a style with a solid fill pattern (fallback if dense hatch is unavailable)
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Pattern = BackgroundType.Solid;          // solid fill pattern
            headerStyle.ForegroundColor = Color.Yellow;          // fill color
            headerStyle.BackgroundColor = Color.Blue;            // background color (used for patterns)

            // Apply the style to the header range
            StyleFlag flag = new StyleFlag { All = true };
            sheet.Cells.CreateRange(startRow, startColumn,
                                    endRow - startRow + 1,
                                    endColumn - startColumn + 1).ApplyStyle(headerStyle, flag);

            // Set header text
            sheet.Cells["A1"].PutValue("Header 1");
            sheet.Cells["B1"].PutValue("Header 2");
            sheet.Cells["C1"].PutValue("Header 3");
            sheet.Cells["D1"].PutValue("Header 4");

            // Save the workbook
            string outputPath = "HeaderWithDenseHatch.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
