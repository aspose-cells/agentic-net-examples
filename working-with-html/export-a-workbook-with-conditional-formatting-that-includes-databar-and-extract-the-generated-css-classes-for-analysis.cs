// Title: Export a workbook with DataBar conditional formatting to HTML and extract the generated CSS classes using Aspose.Cells for .NET (C#)
// Common Searches: how to export a workbook with DataBar conditional formatting to HTML using Aspose.Cells in C# | extract CSS class definitions from HTML saved by Aspose.Cells .NET | Aspose.Cells C# export conditional formatting DataBar to HTML and parse the style block | read generated CSS from Aspose.Cells HTML output for DataBar analysis | save Excel workbook as HTML with DataBar and retrieve CSS classes programmatically
// Tags: Aspose.Cells HTML export with DataBar formatting | C# retrieve CSS from Aspose.Cells generated HTML | parse style block using regex in Aspose.Cells output | conditional formatting DataBar to HTML conversion | save Excel workbook as HTML and analyze CSS

using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using Aspose.Cells;
using System.Drawing;

namespace AsposeCellsDataBarCssExport
{
    // The example creates a workbook, fills column A with values 1‑10, adds a DataBar conditional formatting rule to the range A1:A10, and saves the workbook as an HTML file using Aspose.Cells HtmlSaveOptions. It then reads the generated HTML, extracts the <style> block with a regular expression, collects all CSS class definitions, prints them to the console, and writes them to a separate .css file for further analysis.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Create a new workbook and populate it with sample data
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Fill column A with values 1..10
                for (int i = 0; i < 10; i++)
                {
                    cells[i, 0].PutValue(i + 1);
                }

                // -----------------------------------------------------------------
                // 2. Add Conditional Formatting with a DataBar
                // -----------------------------------------------------------------
                // Define the range A1:A10
                CellArea area = CellArea.CreateCellArea("A1", "A10");

                // Add a new ConditionalFormatting entry and get its condition collection
                int cfIndex = worksheet.ConditionalFormattings.Add();
                FormatConditionCollection cfCollection = worksheet.ConditionalFormattings[cfIndex];
                cfCollection.AddArea(area);

                // Add a DataBar condition
                int conditionIndex = cfCollection.AddCondition(FormatConditionType.DataBar);
                FormatCondition dataBarCondition = cfCollection[conditionIndex];

                // Configure the DataBar
                dataBarCondition.DataBar.MinLength = 0;          // 0% of the cell width
                dataBarCondition.DataBar.MaxLength = 100;        // 100% of the cell width
                dataBarCondition.DataBar.Color = Color.LightGreen; // Set bar color
                dataBarCondition.DataBar.ShowValue = true;

                // -----------------------------------------------------------------
                // 3. Export the workbook to HTML
                // -----------------------------------------------------------------
                string outputHtmlPath = "DataBarConditionalFormatting.html";

                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    ExportImagesAsBase64 = true,
                    ExportActiveWorksheetOnly = true
                    // ExportConditionalFormatting is not available in this version; default behavior exports it
                };

                // Ensure the directory for the output file exists
                try
                {
                    string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputHtmlPath));
                    if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                    {
                        Directory.CreateDirectory(outputDir);
                    }

                    workbook.Save(outputHtmlPath, htmlOptions);
                    Console.WriteLine($"Workbook saved as HTML to: {outputHtmlPath}");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Failed to save HTML file: {saveEx.Message}");
                    return;
                }

                // -----------------------------------------------------------------
                // 4. Read the generated HTML and extract CSS class definitions
                // -----------------------------------------------------------------
                if (!File.Exists(outputHtmlPath))
                {
                    Console.WriteLine("Generated HTML file not found.");
                    return;
                }

                string htmlContent;
                try
                {
                    htmlContent = File.ReadAllText(outputHtmlPath);
                }
                catch (Exception readEx)
                {
                    Console.WriteLine($"Failed to read HTML file: {readEx.Message}");
                    return;
                }

                // Extract the <style> block
                string styleBlockPattern = @"<style[^>]*>(.*?)</style>";
                Match styleBlockMatch = Regex.Match(htmlContent, styleBlockPattern, RegexOptions.Singleline);
                if (!styleBlockMatch.Success)
                {
                    Console.WriteLine("No <style> block found in the generated HTML.");
                    return;
                }

                string cssContent = styleBlockMatch.Groups[1].Value;

                // Find all class definitions
                string classPattern = @"\.([A-Za-z0-9_-]+)\s*\{[^}]*\}";
                MatchCollection classMatches = Regex.Matches(cssContent, classPattern);

                if (classMatches.Count == 0)
                {
                    Console.WriteLine("No CSS classes were found in the style block.");
                    return;
                }

                List<string> extractedClasses = new List<string>();
                foreach (Match match in classMatches)
                {
                    extractedClasses.Add(match.Value.Trim());
                }

                // -----------------------------------------------------------------
                // 5. Output the extracted CSS classes for analysis
                // -----------------------------------------------------------------
                Console.WriteLine("\n--- Extracted CSS Classes ---\n");
                foreach (string cssClass in extractedClasses)
                {
                    Console.WriteLine(cssClass);
                    Console.WriteLine();
                }

                // Optionally, write the extracted CSS to a separate file
                string cssOutputPath = "ExtractedDataBarCss.css";
                try
                {
                    File.WriteAllLines(cssOutputPath, extractedClasses);
                    Console.WriteLine($"Extracted CSS classes have been written to: {cssOutputPath}");
                }
                catch (Exception writeEx)
                {
                    Console.WriteLine($"Failed to write CSS file: {writeEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
