using System;
using System.IO;
using Aspose.Cells;
using System.Drawing;

// Alias to avoid conflict with System.Range (C# 8+)
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsCustomStyleHtmlDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // ---------- Create and apply a custom style ----------
                // Create a style object (lifecycle rule: CreateStyle)
                Style customStyle = workbook.CreateStyle();
                customStyle.Font.IsBold = true;
                customStyle.Font.Color = Color.Red;
                customStyle.ForegroundColor = Color.Yellow;
                customStyle.Pattern = BackgroundType.Solid;

                // Apply the custom style to a range A1:C3 (lifecycle rule: SetStyle)
                AsposeRange styledRange = sheet.Cells.CreateRange("A1", "C3");
                styledRange.SetStyle(customStyle);

                // ---------- Create an unused style (will not be applied) ----------
                Style unusedStyle = workbook.CreateStyle();
                unusedStyle.ForegroundColor = Color.Red; // Distinct color to detect in HTML
                unusedStyle.Pattern = BackgroundType.Solid;
                // Note: unusedStyle is intentionally NOT applied to any cell/range.

                // ---------- Configure HTML save options ----------
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    // Enable exclusion of unused styles (feature rule)
                    ExcludeUnusedStyles = true
                };

                // Save the workbook as HTML (lifecycle rule: save)
                string htmlPath = "StyledOutput.html";
                workbook.Save(htmlPath, saveOptions);

                // ---------- Verify that only used styles appear in the generated HTML ----------
                // Ensure the HTML file exists before reading
                if (File.Exists(htmlPath))
                {
                    // Read the generated HTML file
                    string htmlContent = File.ReadAllText(htmlPath);

                    // The unused style used a red background; check that this color does NOT appear.
                    bool unusedStyleFound = htmlContent
                        .IndexOf("background-color: red", StringComparison.OrdinalIgnoreCase) >= 0;

                    Console.WriteLine("Unused style detected in HTML: " + unusedStyleFound);
                    Console.WriteLine("Verification passed (no unused styles): " + (!unusedStyleFound));
                }
                else
                {
                    Console.WriteLine($"Error: HTML file '{htmlPath}' was not created.");
                }
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}