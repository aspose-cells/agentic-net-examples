// Title: Apply a Bold Yellow Style to Cells E2‑E10 with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, defines a style with a bold font and solid yellow fill, selects the range E2:E10 on the first worksheet, applies the style, and saves the file as CustomStyle_E2_E10.xlsx.
// Keywords: Aspose.Cells | C# | custom style | bold font | yellow fill | range E2:E10 | apply style to cells | Excel formatting | Workbook | Worksheet | SetStyle | CreateStyle
// Common Searches: Aspose.Cells set bold font and yellow background for a range | C# apply custom style to cells E2 to E10 | How to create and reuse a style in Aspose.Cells | Set solid fill color in Aspose.Cells .NET | Apply formatting to a specific cell range using Aspose.Cells
// Developer Intent: Create a style with bold text and yellow background and apply it to the range E2:E10.
// Use Cases: Highlight header rows in financial reports with a bold yellow style. | Emphasize key metrics in dashboards for quick visual identification. | Prepare an Excel template where designated columns are pre‑styled for data entry.
// AI Prompts: Write C# code using Aspose.Cells to define a style with bold font and yellow fill and apply it to a specified range. | Show how to clone and reuse a custom style across multiple worksheets in Aspose.Cells. | Demonstrate conditional styling based on cell values with Aspose.Cells in C#.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsCustomStyleDemo
{
    // Creates a workbook, defines a style with a bold font and solid yellow fill, selects the range E2:E10 on the first worksheet, applies the style, and saves the file as CustomStyle_E2_E10.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Create a custom style
                Style customStyle = workbook.CreateStyle();

                // Set bold font
                customStyle.Font.IsBold = true;

                // Set solid yellow background
                customStyle.Pattern = BackgroundType.Solid;
                customStyle.ForegroundColor = Color.Yellow;

                // Define the target range E2:E10 (use Aspose.Cells.Range to avoid ambiguity with System.Range)
                Aspose.Cells.Range targetRange = worksheet.Cells.CreateRange("E2", "E10");

                // Apply the custom style to the range
                targetRange.SetStyle(customStyle);

                // Save the workbook
                workbook.Save("CustomStyle_E2_E10.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
