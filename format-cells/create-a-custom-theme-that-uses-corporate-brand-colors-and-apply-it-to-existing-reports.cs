// Title: Apply a Custom Corporate Theme with Brand Colors to Multiple Excel Workbooks using Aspose.Cells for .NET (C#)
// Description: This example shows how to define a 12‑color corporate palette, create a reusable custom theme named "CorporateTheme" with Aspose.Cells, apply it to a list of existing Excel files, format a header cell using the Accent1 theme color, and save each workbook while handling missing files and runtime errors.
// Keywords: Aspose.Cells custom theme | C# Excel theme programmatically | brand colors Excel workbook | batch apply Excel theme .NET | Workbook.CustomTheme | ThemeColor Accent1 | Excel report styling | error handling Aspose.Cells
// Common Searches: how to create a custom Excel theme with Aspose.Cells | apply the same theme to multiple workbooks C# | set cell font to Accent1 theme color Aspose.Cells | batch process Excel files with corporate colors | Aspose.Cells custom theme example
// Developer Intent: Create a corporate color palette, build a reusable custom theme, and programmatically apply it to several existing Excel reports.
// Use Cases: Generate a 12‑color corporate theme once and reuse it across any workbook. | Load each report file, attach the custom theme, and style a header cell (A1) with Accent1, bold, size 14. | Iterate over an array of file paths, apply the theme, save the workbook, and gracefully skip or log missing files.
// AI Prompts: Write C# code that builds a custom theme from an array of 12 System.Drawing.Color values and applies it to a workbook with Aspose.Cells. | Show how to set a cell's font to the Accent1 theme color, make it bold, and set the font size to 14 using Aspose.Cells. | Provide a loop that processes a list of Excel file paths, applies a custom corporate theme, updates a specific cell, and saves each workbook with error handling for missing files.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace CorporateThemeDemo
{
    // This example shows how to define a 12‑color corporate palette, create a reusable custom theme named "CorporateTheme" with Aspose.Cells, apply it to a list of existing Excel files, format a header cell using the Accent1 theme color, and save each workbook while handling missing files and runtime errors.
    class Program
    {
        static void Main()
        {
            // Define corporate brand colors (exactly 12 colors)
            Color[] corporateColors = new Color[]
            {
                Color.FromArgb(0, 70, 127),   // Background1
                Color.FromArgb(255, 255, 255),// Text1
                Color.FromArgb(230, 230, 230),// Background2
                Color.FromArgb(0, 0, 0),      // Text2
                Color.FromArgb(0, 112, 192),  // Accent1
                Color.FromArgb(255, 192, 0),  // Accent2
                Color.FromArgb(112, 48, 160), // Accent3
                Color.FromArgb(0, 176, 80),   // Accent4
                Color.FromArgb(255, 0, 0),    // Accent5
                Color.FromArgb(255, 0, 255),  // Accent6
                Color.FromArgb(0, 0, 255),    // Hyperlink
                Color.FromArgb(128, 0, 128)   // Followed Hyperlink
            };

            // List of existing report files to which the theme will be applied
            string[] reportFiles = new string[]
            {
                "ReportQ1.xlsx",
                "ReportQ2.xlsx",
                "ReportQ3.xlsx"
            };

            foreach (string filePath in reportFiles)
            {
                try
                {
                    // Verify that the file exists before attempting to load it
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found: {filePath}. Skipping.");
                        continue;
                    }

                    // Load the existing workbook
                    Workbook workbook = new Workbook(filePath);

                    // Apply the custom corporate theme
                    workbook.CustomTheme("CorporateTheme", corporateColors);

                    // Example: set a cell style to use the new Accent1 theme color
                    Worksheet sheet = workbook.Worksheets[0];
                    Cell demoCell = sheet.Cells["A1"];
                    demoCell.PutValue("Corporate Themed Header");

                    Style style = workbook.CreateStyle();
                    style.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent1, 0.0);
                    style.Font.IsBold = true;
                    style.Font.Size = 14;
                    demoCell.SetStyle(style);

                    // Save the workbook (overwrites the original file)
                    workbook.Save(filePath);

                    Console.WriteLine($"Theme applied and saved: {filePath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Processing completed.");
        }
    }
}
