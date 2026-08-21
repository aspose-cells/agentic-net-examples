// Title: Apply a Corporate Brand Custom Theme to Multiple Excel Reports with Aspose.Cells for .NET (C#)
// Description: Defines a 12‑color corporate palette, loads each workbook (creates a new one if missing), uses Workbook.CustomTheme to set the brand theme, optionally formats a demo cell with an Accent1 font, and saves the themed workbook as a new file.
// Keywords: Aspose.Cells | custom theme | brand colors | C# | .NET | Workbook.CustomTheme | Excel theme programmatically | apply theme to multiple workbooks | theme colors array | Excel branding
// Common Searches: Aspose.Cells create custom Excel theme C# | apply corporate color palette to existing workbooks Aspose.Cells | set theme colors programmatically .NET | bulk apply Excel theme to multiple files using Aspose.Cells | use ThemeColor in cell style after custom theme Aspose.Cells
// Developer Intent: Programmatically define a corporate color palette and apply it as a custom Excel theme to one or more existing workbooks using Aspose.Cells for .NET.
// Use Cases: Standardize branding across quarterly financial reports by applying the same custom theme to each workbook. | Generate themed copies of legacy reports without altering the original files. | Create new workbooks with corporate branding when source files are missing. | Demonstrate theme color usage by formatting a cell with an Accent1 font after the theme is applied.
// AI Prompts: Generate C# code that creates a 12‑color custom theme and applies it to a collection of Excel files using Aspose.Cells. | Show how to retrieve the current theme colors from a workbook after applying a custom theme with Aspose.Cells. | Explain how to use ThemeColor (Accent1, Accent2, etc.) in cell styles after setting a custom theme in Aspose.Cells for .NET. | Provide a script to batch‑process a folder of .xlsx files, apply a corporate theme, and save them with a suffix.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace CorporateThemeDemo
{
    // Defines a 12‑color corporate palette, loads each workbook (creates a new one if missing), uses Workbook.CustomTheme to set the brand theme, optionally formats a demo cell with an Accent1 font, and saves the themed workbook as a new file.
    public class ApplyCustomTheme
    {
        // Define corporate brand colors (12 theme colors)
        private static Color[] GetCorporateColors()
        {
            // Example corporate colors – replace with actual brand colors
            return new Color[]
            {
                Color.FromArgb(255, 255, 255), // Background1 (white)
                Color.FromArgb(0, 0, 0),       // Text1 (black)
                Color.FromArgb(240, 240, 240), // Background2 (light gray)
                Color.FromArgb(80, 80, 80),    // Text2 (dark gray)
                Color.FromArgb(0, 120, 215),   // Accent1 (primary brand blue)
                Color.FromArgb(0, 153, 0),     // Accent2 (secondary brand green)
                Color.FromArgb(255, 153, 0),   // Accent3 (accent orange)
                Color.FromArgb(255, 0, 0),     // Accent4 (accent red)
                Color.FromArgb(128, 0, 128),   // Accent5 (accent purple)
                Color.FromArgb(255, 255, 0),   // Accent6 (accent yellow)
                Color.FromArgb(0, 0, 255),     // Hyperlink (blue)
                Color.FromArgb(128, 0, 0)      // Followed Hyperlink (maroon)
            };
        }

        public static void Run()
        {
            // List of existing report file paths to which the theme will be applied
            string[] reportFiles = new string[]
            {
                "Report1.xlsx",
                "Report2.xlsx",
                // Add more report file names as needed
            };

            // Prepare the custom theme colors
            Color[] corporateColors = GetCorporateColors();

            foreach (string filePath in reportFiles)
            {
                try
                {
                    Workbook workbook;

                    // Load the existing workbook if it exists; otherwise create a new one
                    if (File.Exists(filePath))
                    {
                        workbook = new Workbook(filePath);
                    }
                    else
                    {
                        Console.WriteLine($"File '{filePath}' not found. Creating a new workbook.");
                        workbook = new Workbook();
                        workbook.Worksheets[0].Name = "Sheet1";
                    }

                    // Apply the custom corporate theme
                    workbook.CustomTheme("CorporateBrandTheme", corporateColors);

                    // OPTIONAL: Demonstrate theme usage by setting a styled cell
                    Worksheet sheet = workbook.Worksheets[0];
                    Cell demoCell = sheet.Cells["A1"];
                    demoCell.PutValue("Corporate Themed Report");
                    Style demoStyle = workbook.CreateStyle();
                    demoStyle.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent1, 0.0);
                    demoStyle.Font.IsBold = true;
                    demoStyle.Font.Size = 14;
                    demoCell.SetStyle(demoStyle);

                    // Save the workbook (overwrites the original file or saves as new)
                    string outputPath = Path.GetFileNameWithoutExtension(filePath) + "_Themed.xlsx";
                    workbook.Save(outputPath);
                    Console.WriteLine($"Applied corporate theme to '{filePath}' and saved as '{outputPath}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                ApplyCustomTheme.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
