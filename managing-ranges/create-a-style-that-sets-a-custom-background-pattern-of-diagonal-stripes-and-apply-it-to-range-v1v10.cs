// Title: Apply a Diagonal Stripe Background Style to Range V1:V10 with Aspose.Cells for .NET (C#)
// Description: Shows how to create a Style with a DiagonalStripe pattern, set its foreground and background colors, define the V1:V10 range, apply the style, and save the workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | diagonal stripe pattern | cell background style | range V1:V10 | CreateStyle | SetStyle | Workbook saving | custom cell formatting
// Common Searches: Aspose.Cells set diagonal stripe pattern | C# apply background pattern to column | How to style cells V1 to V10 Aspose.Cells | Create custom cell style Aspose.Cells .NET | Set foreground and background colors Aspose.Cells
// Developer Intent: Create a diagonal‑stripe style and apply it to cells V1 through V10.
// Use Cases: Visually separate a column in an exported report with a striped background. | Highlight cells that satisfy a business rule using a diagonal stripe pattern. | Design custom header formatting for generated Excel files. | Reuse the same patterned style across multiple worksheets or non‑contiguous ranges.
// AI Prompts: Write C# code using Aspose.Cells to define a diagonal stripe style with specific foreground and background colors and apply it to a given range. | Explain how to change the pattern type and colors of an existing Aspose.Cells style at runtime. | Show how to apply the same diagonal stripe style to several non‑adjacent ranges in one workbook. | Provide a snippet that saves the workbook after styling without overwriting an existing file.

using Aspose.Cells;
using System;
using System.Drawing;

// Alias to avoid conflict with System.Range
using AsposeRange = Aspose.Cells.Range;

// Shows how to create a Style with a DiagonalStripe pattern, set its foreground and background colors, define the V1:V10 range, apply the style, and save the workbook using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Create a style with diagonal stripe pattern
            Style style = workbook.CreateStyle();
            style.Pattern = BackgroundType.DiagonalStripe;      // set pattern type
            style.ForegroundColor = Color.LightBlue;            // stripe (foreground) color
            style.BackgroundColor = Color.DarkBlue;             // background color

            // Define the range V1:V10 using the aliased AsposeRange
            AsposeRange range = worksheet.Cells.CreateRange("V1", "V10");

            // Apply the style to the entire range
            range.SetStyle(style);

            // Save the workbook
            workbook.Save("DiagonalStripeStyle.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
