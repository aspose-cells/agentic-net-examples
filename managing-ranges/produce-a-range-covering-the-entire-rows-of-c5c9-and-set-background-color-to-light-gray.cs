// Title: C# – Apply Light Gray Background to Range C5:C9 with Aspose.Cells
// Description: Demonstrates how to create a workbook, define the range C5:C9, build a solid light‑gray style, apply it to the range, and save the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | range C5:C9 | set background color | light gray fill | style cells programmatically | CreateRange | SetStyle | save workbook
// Common Searches: Aspose.Cells set background color for a range | C# apply gray fill to cells C5 to C9 | How to style a range in Aspose.Cells .NET | Create and format range C5:C9 using Aspose.Cells | Save workbook after applying cell style Aspose.Cells
// Developer Intent: Add a solid light‑gray fill to cells C5 through C9 in a worksheet.
// Use Cases: Highlight header or subtotal rows in a report. | Visually group a column segment in a generated template. | Improve readability of financial statements by shading specific rows.
// AI Prompts: Show C# code that creates a range C5:C9 and sets a light gray background with Aspose.Cells. | How can I reuse a style to format multiple ranges in Aspose.Cells for .NET? | Explain the steps to apply a solid fill color to a cell range and save the workbook using Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;

// Demonstrates how to create a workbook, define the range C5:C9, build a solid light‑gray style, apply it to the range, and save the file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Create a range that covers cells C5:C9
            Aspose.Cells.Range range = cells.CreateRange("C5", "C9");

            // Define a style with a solid light gray background
            Style style = workbook.CreateStyle();
            style.Pattern = BackgroundType.Solid;
            style.ForegroundColor = Color.LightGray;

            // Apply the style to the range
            range.SetStyle(style);

            // Save the workbook
            workbook.Save("RangeLightGray.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
