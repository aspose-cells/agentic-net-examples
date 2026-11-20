// Title: Apply the workbook's Accent1 theme color as a solid fill to an imported cell range using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set solid background of a range to the workbook Accent1 theme color | How to apply theme Accent1 fill to imported data cells with Aspose.Cells .NET | Retrieve workbook theme color and style a range in Aspose.Cells C# example | Using GetThemeColor and ApplyStyle for a data import range in Aspose.Cells | Formatting imported cells with workbook theme colors in Aspose.Cells for .NET
// Tags: apply theme accent1 fill Aspose.Cells | style range with workbook theme color C# | solid background formatting Aspose.Cells .NET | imported data cell styling Aspose.Cells | GetThemeColor usage Aspose.Cells

using Aspose.Cells;
using System;
using System.Drawing;

// Alias to avoid conflict with System.Range
using AsposeRange = Aspose.Cells.Range;

// The example creates a new workbook, imports sample data into cells A1:C3, retrieves the workbook's Accent1 theme color via GetThemeColor, builds a style with a solid fill using that color, applies the style to the imported range with ApplyStyle, and saves the workbook as Output.xlsx.
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

            // ----- Data import routine (example) -----
            // Populate cells A1:C3 with sample data
            string[,] data = new string[,]
            {
                { "Item", "Quantity", "Price" },
                { "Apple", "10", "0.5" },
                { "Banana", "20", "0.3" }
            };

            for (int r = 0; r < data.GetLength(0); r++)
            {
                for (int c = 0; c < data.GetLength(1); c++)
                {
                    sheet.Cells[r, c].PutValue(data[r, c]);
                }
            }

            // ----- Apply theme's Accent1 color to the fill of imported cells -----
            // Retrieve the Accent1 color from the workbook's theme
            Color accent1Color = workbook.GetThemeColor(ThemeColorType.Accent1);

            // Create a style with solid fill using Accent1
            Style accentStyle = workbook.CreateStyle();
            accentStyle.ForegroundColor = accent1Color;
            accentStyle.Pattern = BackgroundType.Solid;

            // Define a style flag to apply all style attributes
            StyleFlag flag = new StyleFlag { All = true };

            // Apply the style to the range that contains the imported data (A1:C3)
            AsposeRange importedRange = sheet.Cells.CreateRange("A1:C3");
            importedRange.ApplyStyle(accentStyle, flag);

            // Save the workbook to a file
            workbook.Save("Output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
