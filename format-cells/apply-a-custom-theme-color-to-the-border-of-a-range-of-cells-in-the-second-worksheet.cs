// Title: Apply a custom teal theme color to the borders of a cell range in the second worksheet using Aspose.Cells for .NET
// AI Prompts: Create a thin teal border around cells B2:D5 on the second worksheet and save the workbook as Output.xlsx. | Instantiate a System.Drawing.Color and set it as the border color for all sides of a range using an Aspose.Cells style object. | Apply a style with thin borders and a custom color to a specific range by configuring a StyleFlag to apply all style elements.
// Common Searches: Aspose.Cells C# set a unique border hue for a range on a specific sheet | How to apply a theme color to cell borders on the second sheet with Aspose.Cells .NET | Using StyleFlag to apply border formatting to a range in Aspose.Cells for .NET | Apply teal border to B2:D5 in an Aspose.Cells workbook
// Tags: border color customization Aspose.Cells .NET | apply style to worksheet range Aspose.Cells | thin border line style C# | StyleFlag All=true border style application | color object for Excel styling Aspose.Cells

using System;
using Aspose.Cells;
using System.Drawing;

// // This program creates a new workbook, adds a second worksheet, defines the range B2:D5, creates a style with thin borders, sets a custom teal color for all four borders, applies the style to the range using a StyleFlag, and saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Add a second worksheet (index 1)
            workbook.Worksheets.Add("SecondSheet");
            var sheet = workbook.Worksheets[1];

            // Define the range whose borders will be styled (e.g., B2:D5)
            var range = sheet.Cells.CreateRange("B2:D5");

            // Create a new style object
            var style = workbook.CreateStyle();

            // Set thin border lines for all four sides
            style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;

            // Define a custom theme color (example: teal)
            var customColor = Color.FromArgb(0, 128, 128);

            // Apply the custom color to each border side
            style.Borders[BorderType.LeftBorder].Color = customColor;
            style.Borders[BorderType.RightBorder].Color = customColor;
            style.Borders[BorderType.TopBorder].Color = customColor;
            style.Borders[BorderType.BottomBorder].Color = customColor;

            // Define which style elements to apply (all in this case)
            var styleFlag = new StyleFlag { All = true };

            // Apply the modified style to the range
            range.ApplyStyle(style, styleFlag);

            // Save the workbook to a file
            workbook.Save("Output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
