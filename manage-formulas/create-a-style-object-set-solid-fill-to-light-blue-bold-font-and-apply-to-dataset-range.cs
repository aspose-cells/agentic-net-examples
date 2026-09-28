// Title: Create a light‑blue solid fill style with bold font and apply it to a cell range using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that defines a Style having a solid light‑blue background and bold text, then applies it to the range A1:C10. | Demonstrate how to use a StyleFlag with All=true to apply every style attribute to a specified range in Aspose.Cells. | Adapt the example to use a different background color (e.g., LightGreen) and a different cell range while preserving the bold font setting.
// Common Searches: Aspose.Cells C# set solid fill color and bold font for a cell range | apply custom style to A1:C10 using Aspose.Cells .NET | how to use StyleFlag to apply full style in Aspose.Cells | change background color of a range in Aspose.Cells workbook C# | save styled worksheet as Excel file with Aspose.Cells
// Tags: Aspose.Cells solid background style | C# apply style to cell range Aspose.Cells | Aspose.Cells StyleFlag usage | Aspose.Cells bold font formatting | Aspose.Cells save styled workbook

using System;
using System.Drawing;
using Aspose.Cells;
using CellsRange = Aspose.Cells.Range;

// The program creates a new workbook, defines the range A1:C10, builds a Style with a solid light‑blue fill and bold font, applies the style to the range using a StyleFlag with All=true, and saves the workbook as StyledDataSet.xlsx.
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

            // Define the range to which the style will be applied (example: A1:C10)
            CellsRange dataRange = sheet.Cells.CreateRange("A1", "C10");

            // Create a new style object
            Style style = workbook.CreateStyle();

            // Set solid fill to light blue
            style.ForegroundColor = Color.LightBlue;
            style.Pattern = BackgroundType.Solid;

            // Set the font to bold
            style.Font.IsBold = true;

            // Apply the style to the entire range
            StyleFlag flag = new StyleFlag
            {
                All = true // Apply all style attributes
            };
            dataRange.ApplyStyle(style, flag);

            // Save the workbook with the applied style
            workbook.Save("StyledDataSet.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
