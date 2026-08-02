// Title: Apply a Diagonal Stripe Background Style to Cells V1‑V10 with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a Style with BackgroundType.DiagonalStripe, set foreground and background colors, target the range V1:V10 on the first worksheet, apply the style, and save the workbook as DiagonalStripeStyle.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | diagonal stripe style | background pattern | range styling | Excel workbook | Style.Create | BackgroundType.DiagonalStripe | cell formatting | V1:V10
// Common Searches: Aspose.Cells set diagonal stripe background | C# apply pattern style to range V1 V10 | How to use BackgroundType.DiagonalStripe in .NET | Styling a column with striped background in Aspose.Cells | Save workbook after applying custom style Aspose.Cells
// Developer Intent: Create a diagonal‑stripe style and apply it to the V1:V10 range in a .NET workbook.
// Use Cases: Visually separate a column in a report with a striped background. | Highlight cells that require attention using a custom pattern. | Design header rows with diagonal stripes for branding consistency.
// AI Prompts: Generate C# code that defines a diagonal stripe style and applies it to a specified range with Aspose.Cells. | Explain how to modify the foreground and background colors of a diagonal stripe pattern in Aspose.Cells for .NET. | Show how to reuse a previously created style across multiple worksheets in the same workbook.

using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Demonstrates how to create a Style with BackgroundType.DiagonalStripe, set foreground and background colors, target the range V1:V10 on the first worksheet, apply the style, and save the workbook as DiagonalStripeStyle.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Create a style with diagonal stripe background pattern
            Style style = workbook.CreateStyle();
            style.Pattern = BackgroundType.DiagonalStripe;   // set pattern type
            style.ForegroundColor = Color.LightBlue;          // color of the stripes
            style.BackgroundColor = Color.DarkBlue;           // background color behind the stripes

            // Define the target range V1:V10
            AsposeRange range = worksheet.Cells.CreateRange("V1", "V10");

            // Apply the style to the entire range
            range.SetStyle(style);

            // Save the workbook
            workbook.Save("DiagonalStripeStyle.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
