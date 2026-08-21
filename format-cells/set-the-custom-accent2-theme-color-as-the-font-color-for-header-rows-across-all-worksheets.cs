// Title: Set Accent2 Theme Font Color for Header Rows in All Worksheets – Aspose.Cells C# Example
// Description: Shows how to define a custom Accent2 theme color, create a style that uses it for the font, and apply that style to the first row of every worksheet in a workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | Excel | theme color | Accent2 | header row style | StyleFlag | FontColor | multi‑sheet workbook | programmatic formatting
// Common Searches: Aspose.Cells set Accent2 theme color header | C# apply theme font color to all worksheet headers | How to use StyleFlag FontColor in Aspose.Cells | Apply custom theme color to Excel header rows .NET | Set theme color for multiple sheets Aspose.Cells
// Developer Intent: Apply a custom Accent2 theme color as the font color for header rows across every worksheet in an Excel workbook.
// Use Cases: Generate a multi‑sheet report where each sheet’s header row automatically follows the Accent2 theme for consistent branding. | Programmatically create workbooks and apply a theme‑based font color to header rows without manual cell formatting. | Change the Accent2 theme color once and have all header rows in the workbook update instantly.
// AI Prompts: Write C# code with Aspose.Cells that sets a custom Accent2 theme color and applies it as the font color to the first row of every worksheet. | Provide an Aspose.Cells .NET example that creates a style using ThemeColor Accent2, applies it to header rows across all sheets, and saves the workbook. | Explain how to modify the sample to target only selected worksheets instead of all worksheets when applying the Accent2 theme font color.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// Alias to avoid conflict with System.Range
using Range = Aspose.Cells.Range;

namespace AsposeCellsHeaderTheme
{
    // Shows how to define a custom Accent2 theme color, create a style that uses it for the font, and apply that style to the first row of every worksheet in a workbook using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // First worksheet (default sheet)
                Worksheet ws1 = workbook.Worksheets[0];
                ws1.Name = "Sheet1";
                ws1.Cells["A1"].PutValue("Header1");
                ws1.Cells["B1"].PutValue("Header2");
                ws1.Cells["A2"].PutValue("Data1");
                ws1.Cells["B2"].PutValue("Data2");

                // Add a second worksheet (Worksheets.Add() returns the index in some versions)
                int ws2Index = workbook.Worksheets.Add();
                Worksheet ws2 = workbook.Worksheets[ws2Index];
                ws2.Name = "Sheet2";
                ws2.Cells["A1"].PutValue("HeaderA");
                ws2.Cells["B1"].PutValue("HeaderB");
                ws2.Cells["A2"].PutValue("Info1");
                ws2.Cells["B2"].PutValue("Info2");

                // Set a custom Accent2 theme color
                workbook.SetThemeColor(ThemeColorType.Accent2, Color.Teal);

                // Prepare a style that uses the Accent2 theme color for the font
                Style headerStyle = workbook.CreateStyle();
                headerStyle.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent2, 0);

                // Apply only the font color
                StyleFlag flag = new StyleFlag { FontColor = true };

                // Apply the style to the first row of every worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    int usedColumns = sheet.Cells.MaxColumn + 1; // number of used columns
                    Range headerRange = sheet.Cells.CreateRange(0, 0, 1, usedColumns);
                    headerRange.ApplyStyle(headerStyle, flag);
                }

                // Define output path and ensure the directory exists
                string outputPath = "HeaderWithAccent2Theme.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? string.Empty;
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
