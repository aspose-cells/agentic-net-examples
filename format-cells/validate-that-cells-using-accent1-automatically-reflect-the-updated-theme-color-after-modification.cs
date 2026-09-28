// Title: Validate that a cell’s solid fill foreground color matches the assigned RGB value using Aspose.Cells for .NET
// AI Prompts: Create a workbook, apply a solid red fill to cell A1 via a style, read the cell’s style back, and assert that the ForegroundColor equals Color.Red. | Adapt the sample to use the Accent1 theme color instead of a direct RGB value, then change the workbook’s theme and confirm the cell updates automatically. | Loop through a range of cells, assign each a different solid fill color, and programmatically verify that each cell’s ForegroundColor property matches the expected color.
// Common Searches: Aspose.Cells C# how to verify cell background color after applying a style | read ForegroundColor of a styled Excel cell using Aspose.Cells .NET | check if Excel cell updates when theme accent color changes with Aspose.Cells | C# code to assert solid fill color of a specific cell in a generated workbook
// Tags: solid fill foreground color verification Aspose.Cells | retrieve cell style ForegroundColor C# | theme accent propagation validation Aspose.Cells | assert cell background color .NET Excel library | programmatic style comparison Aspose.Cells workbook

using System;
using System.Drawing;
using Aspose.Cells;

// The example creates a new workbook, applies a solid red fill to cell A1 using a style, reads back the cell’s style to confirm the ForegroundColor matches Color.Red, prints the validation result, and saves the file as ThemeAccentValidation.xlsx.
class ThemeAccentValidation
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Set cell A1 to a solid fill with a specific color (Red)
            Cell cell = sheet.Cells["A1"];
            Style style = workbook.CreateStyle();
            style.ForegroundColor = Color.Red;          // Directly set the color
            style.Pattern = BackgroundType.Solid;       // Solid fill
            cell.SetStyle(style);

            // Retrieve the cell's style after setting the color
            Style updatedStyle = cell.GetStyle();

            // The ForegroundColor property reflects the actual color set
            Color actualColor = updatedStyle.ForegroundColor;

            // Validate that the cell reflects the expected color (Red)
            if (actualColor.ToArgb() == Color.Red.ToArgb())
            {
                Console.WriteLine("Validation passed: Cell reflects the expected color.");
            }
            else
            {
                Console.WriteLine("Validation failed: Cell does not reflect the expected color.");
            }

            // Save the workbook (optional) – ensure the directory exists
            string outputPath = "ThemeAccentValidation.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
