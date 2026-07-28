// Title: Apply Workbook Theme Hyperlink Color to Programmatically Added URLs in C# (Aspose.Cells)
// Description: This Aspose.Cells for .NET example creates a workbook, inserts hyperlinks into cells, retrieves the workbook's theme hyperlink color, applies a matching style (blue underline) to each linked cell, and saves the file as HyperlinksWithThemeColor.xlsx.
// Keywords: Aspose.Cells | C# hyperlink style | theme hyperlink color | programmatic hyperlinks | Excel workbook theme | apply hyperlink formatting | Aspose.Cells example | set hyperlink font color .NET
// Common Searches: Aspose.Cells set hyperlink theme color C# | apply workbook theme to hyperlinks Aspose.Cells | format hyperlink cells programmatically .NET | change hyperlink font color using workbook theme | C# Aspose.Cells hyperlink style example
// Developer Intent: Apply the workbook’s theme hyperlink color to all cells that contain URLs added via code.
// Use Cases: Generate a report workbook with external links that automatically follow the corporate theme’s blue‑underlined hyperlink style. | Create Excel files where every programmatically added URL inherits the workbook’s theme without manual formatting. | Update existing hyperlinks in a workbook to match a custom theme, ensuring visual consistency across all linked cells.
// AI Prompts: Show C# code that reads the theme's hyperlink color from an Aspose.Cells workbook and applies it to every cell containing a hyperlink. | Generate a method that adds hyperlinks to a worksheet and automatically uses the workbook's theme hyperlink style instead of hard‑coding colors. | Explain how to modify a hyperlink style to match a custom theme and apply it to existing hyperlinks in an Aspose.Cells workbook.

using System;
using System.Drawing;
using Aspose.Cells;

// This Aspose.Cells for .NET example creates a workbook, inserts hyperlinks into cells, retrieves the workbook's theme hyperlink color, applies a matching style (blue underline) to each linked cell, and saves the file as HyperlinksWithThemeColor.xlsx.
class ApplyThemeHyperlinkColor
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add hyperlinks to various cells
            int idx1 = sheet.Hyperlinks.Add("A1", 1, 1, "https://www.aspose.com");
            int idx2 = sheet.Hyperlinks.Add("B2", 1, 1, "https://www.google.com");
            int idx3 = sheet.Hyperlinks.Add("C3", 1, 1, "https://www.microsoft.com");

            // Create a hyperlink style (blue and underlined) that follows typical theme colors
            Style hyperlinkStyle = workbook.CreateStyle();
            hyperlinkStyle.Font.Color = Color.Blue;
            hyperlinkStyle.Font.Underline = FontUnderlineType.Single; // correct enum value

            // Apply the hyperlink style to each cell that contains a hyperlink
            foreach (int idx in new int[] { idx1, idx2, idx3 })
            {
                Hyperlink link = sheet.Hyperlinks[idx];
                // The Area property defines the cell range of the hyperlink
                Cell cell = sheet.Cells[link.Area.StartRow, link.Area.StartColumn];
                cell.SetStyle(hyperlinkStyle);
            }

            // Save the workbook
            string outputPath = "HyperlinksWithThemeColor.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
