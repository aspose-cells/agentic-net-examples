// Title: Aspose.Cells .NET – Apply Bold Font and Yellow Background Style to Range E2:E10
// Description: Create a custom Style with a bold typeface and solid yellow fill, then apply it to cells E2 through E10 in a new workbook using Aspose.Cells for .NET and save as StyledRange.xlsx.
// Keywords: Aspose.Cells C# style | custom cell style Aspose.Cells | bold font yellow background | apply style to range | E2:E10 formatting | Aspose.Cells .NET example | Workbook styling C# | range SetStyle Aspose
// Common Searches: how to set bold and yellow style for a range in Aspose.Cells | Aspose.Cells apply custom style to multiple cells | C# Aspose.Cells set background color for column range | create and reuse style in Aspose.Cells workbook
// Developer Intent: Define a bold, yellow‑filled style and apply it to the cell range E2:E10 in a workbook.
// Use Cases: Highlight a column of key metrics in a financial dashboard. | Mark required input cells in a data‑entry template. | Create a reusable header style for generated reports.
// AI Prompts: Write C# code with Aspose.Cells that creates a bold, yellow style and applies it to E2:E10. | Show how to store the custom style in a variable and reuse it on other ranges in the same workbook. | Explain how to add a thin black border to the existing bold‑yellow style while keeping the same range application.

using System;
using Aspose.Cells;
using System.Drawing;

// Create a custom Style with a bold typeface and solid yellow fill, then apply it to cells E2 through E10 in a new workbook using Aspose.Cells for .NET and save as StyledRange.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Create a custom style: bold font and yellow background
            Style customStyle = workbook.CreateStyle();
            customStyle.Font.IsBold = true;                     // Bold font
            customStyle.Pattern = BackgroundType.Solid;         // Enable solid fill
            customStyle.ForegroundColor = Color.Yellow;         // Yellow background

            // Define the target range E2:E10 (use Aspose.Cells.Range to avoid conflict with System.Range)
            Aspose.Cells.Range targetRange = worksheet.Cells.CreateRange("E2", "E10");

            // Apply the custom style to the entire range
            targetRange.SetStyle(customStyle);

            // Save the workbook to a file
            workbook.Save("StyledRange.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
