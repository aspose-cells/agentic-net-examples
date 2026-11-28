// Title: Export Aspose.Cells workbook to HTML with separate CSS and check Firefox border rendering (C#)
// Common Searches: how to export Excel workbook to HTML with separate CSS file using Aspose.Cells C# | Aspose.Cells C# generate HTML and CSS for custom cell borders | verify Firefox border rendering of HTML exported from Aspose.Cells | inspect generated CSS for cell borders after Aspose.Cells HTML save | apply different border styles to a range and export to HTML with Aspose.Cells C#
// Tags: Aspose.Cells HTML export with custom borders | C# separate CSS generation from workbook | Firefox CSS border fallback testing | Aspose.Cells range style application | HtmlSaveOptions ExportActiveWorksheetOnly usage

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;
using System.Drawing;

namespace AsposeCellsBorderFallbackTest
{
    // The example creates a workbook, applies distinct top, bottom, left, and right borders with various styles and colors to a range, and saves the workbook as HTML using HtmlSaveOptions that output CSS to a separate file. It then reads and prints the generated CSS (or embedded style block) and provides console instructions for manually inspecting the rendered borders in Firefox to confirm proper fallback behavior.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Prepare output folder
                // -----------------------------------------------------------------
                string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "output");
                Directory.CreateDirectory(outputFolder);

                // -----------------------------------------------------------------
                // 2. Create a workbook with some styled cells
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Fill sample data
                sheet.Cells["A1"].PutValue("Header");
                sheet.Cells["B2"].PutValue("Cell with borders");
                sheet.Cells["C3"].PutValue("Another cell");

                // Create a style with distinct borders
                Style borderStyle = workbook.CreateStyle();
                borderStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thick;
                borderStyle.Borders[BorderType.TopBorder].Color = Color.Red;

                borderStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Dashed;
                borderStyle.Borders[BorderType.BottomBorder].Color = Color.Green;

                borderStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Dotted;
                borderStyle.Borders[BorderType.LeftBorder].Color = Color.Blue;

                borderStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Double;
                borderStyle.Borders[BorderType.RightBorder].Color = Color.Purple;

                // Apply the style to a range (use fully qualified Aspose.Cells.Range to avoid ambiguity)
                Aspose.Cells.Range range = sheet.Cells.CreateRange("B2:C3");
                range.ApplyStyle(borderStyle, new StyleFlag { Borders = true });

                // -----------------------------------------------------------------
                // 3. Configure HTML export options
                // -----------------------------------------------------------------
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    ExportActiveWorksheetOnly = true,
                    ExportWorksheetCSSSeparately = true // keep CSS in a separate file
                };

                string htmlFilePath = Path.Combine(outputFolder, "WorkbookWithBorders.html");
                string cssFilePath = Path.Combine(outputFolder, "WorkbookWithBorders.css");

                // -----------------------------------------------------------------
                // 4. Save the workbook as HTML
                // -----------------------------------------------------------------
                workbook.Save(htmlFilePath, htmlOptions);
                Console.WriteLine($"HTML file saved to: {htmlFilePath}");
                Console.WriteLine($"CSS file saved to: {cssFilePath}");

                // -----------------------------------------------------------------
                // 5. Load the generated CSS and output relevant parts
                // -----------------------------------------------------------------
                if (File.Exists(cssFilePath))
                {
                    string cssContent = File.ReadAllText(cssFilePath);
                    Console.WriteLine("\n--- Extracted CSS (first 500 characters) ---\n");
                    Console.WriteLine(cssContent.Length > 500 ? cssContent.Substring(0, 500) + "..." : cssContent);
                }
                else if (File.Exists(htmlFilePath))
                {
                    // If CSS is embedded, extract from the HTML <style> block
                    string htmlContent = File.ReadAllText(htmlFilePath);
                    Match styleMatch = Regex.Match(htmlContent, @"<style[^>]*>(.*?)</style>", RegexOptions.Singleline);
                    if (styleMatch.Success)
                    {
                        string embeddedCss = styleMatch.Groups[1].Value.Trim();
                        Console.WriteLine("\n--- Embedded CSS (first 500 characters) ---\n");
                        Console.WriteLine(embeddedCss.Length > 500 ? embeddedCss.Substring(0, 500) + "..." : embeddedCss);
                    }
                    else
                    {
                        Console.WriteLine("No CSS block found in the generated HTML.");
                    }
                }
                else
                {
                    Console.WriteLine("Neither CSS nor HTML files were generated.");
                }

                // -----------------------------------------------------------------
                // 6. Instructions for manual verification in Firefox
                // -----------------------------------------------------------------
                Console.WriteLine("\n=== Manual Verification Steps ===");
                Console.WriteLine("1. Open the generated HTML file in Firefox:");
                Console.WriteLine($"   file://{htmlFilePath}");
                Console.WriteLine("2. Press F12 to open Developer Tools.");
                Console.WriteLine("3. Inspect the cells (e.g., the cell at B2).");
                Console.WriteLine("4. Observe the computed CSS for borders. The browser should render combined borders where possible, falling back to individual borders when styles differ.");
                Console.WriteLine("5. Compare the rendered borders with the CSS definitions printed above.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during execution:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
