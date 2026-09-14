// Title: Apply a corporate brand color theme to an existing Excel workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that opens an existing .xlsx file, defines five corporate brand colors, assigns them to ThemeColorType.Accent1‑Accent4 and Hyperlink using Workbook.SetThemeColor, and saves the modified workbook. | Show how to add file‑existence verification and try‑catch blocks around the theme‑application logic when using Aspose.Cells. | Demonstrate updating only the theme colors of a workbook without altering cell data or formatting, using the Aspose.Cells SetThemeColor method.
// Common Searches: Aspose.Cells C# set custom theme accent colors in an existing workbook | how to replace Excel theme palette with corporate colors using Aspose.Cells | C# code example for applying a brand color scheme to a .xlsx file with Aspose.Cells | error handling when loading a workbook before applying a custom theme in Aspose.Cells | set hyperlink theme color programmatically with Aspose.Cells .NET
// Tags: set custom theme colors Aspose.Cells | apply corporate palette to Excel workbook C# | Workbook.SetThemeColor usage example | load and save themed .xlsx with Aspose.Cells | file not found handling Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsThemeExample
{
    // Loads an existing Excel file, defines corporate brand colors, assigns them to theme accents and hyperlink via Workbook.SetThemeColor, validates file existence, handles exceptions, and saves the themed workbook as a new .xlsx file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Verify that the source report file exists to avoid FileNotFoundException
                const string sourceFile = "ExistingReport.xlsx";
                if (!File.Exists(sourceFile))
                {
                    throw new FileNotFoundException($"The source workbook '{sourceFile}' was not found.");
                }

                // Define corporate brand colors
                Color corporatePrimary   = Color.FromArgb(0, 112, 192);   // Brand Blue
                Color corporateSecondary = Color.FromArgb(255, 192, 0);   // Brand Orange
                Color corporateAccent1   = Color.FromArgb(0, 176, 80);    // Brand Green
                Color corporateAccent2   = Color.FromArgb(192, 0, 0);     // Brand Red
                Color corporateHyperlink = Color.FromArgb(0, 0, 255);    // Hyperlink Blue

                // Load the existing report workbook
                Workbook reportWorkbook = new Workbook(sourceFile);

                // Apply the custom theme colors to the workbook
                reportWorkbook.SetThemeColor(ThemeColorType.Accent1, corporatePrimary);
                reportWorkbook.SetThemeColor(ThemeColorType.Accent2, corporateSecondary);
                reportWorkbook.SetThemeColor(ThemeColorType.Accent3, corporateAccent1);
                reportWorkbook.SetThemeColor(ThemeColorType.Accent4, corporateAccent2);
                reportWorkbook.SetThemeColor(ThemeColorType.Hyperlink, corporateHyperlink);

                // Save the themed report
                const string outputFile = "ExistingReport_Themed.xlsx";
                reportWorkbook.Save(outputFile);
                Console.WriteLine($"The themed workbook has been saved as '{outputFile}'.");
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.Error.WriteLine($"File error: {fnfEx.Message}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
