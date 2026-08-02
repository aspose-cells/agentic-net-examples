// Title: C# – Export Aspose.Cells workbook to HTML with a data‑cell attribute and verify every <td>
// Description: Creates a workbook, fills it with sample data, sets HtmlSaveOptions.CellNameAttribute to "data-cell", saves the sheet as HTML, reads the output, and uses pattern matching to confirm that each <td> element contains the required attribute. Reports success or failure.
// Keywords: Aspose.Cells HTML export C# | CellNameAttribute data-cell | verify td attribute C# | regex HTML table validation | HtmlAgilityPack cell attribute check | export worksheet to HTML with identifiers
// Common Searches: Aspose.Cells add data-cell attribute to HTML cells | C# verify every <td> has custom attribute after export | how to use HtmlSaveOptions.CellNameAttribute | check HTML table cells for attribute in C#
// Developer Intent: Generate HTML from a spreadsheet and ensure each table cell includes a specific data‑cell attribute for downstream processing or testing.
// Use Cases: Add unique cell identifiers to HTML for client‑side scripts or automated UI tests. | Validate exported HTML before publishing to catch missing attributes. | Integrate attribute verification into CI pipelines to guarantee export consistency.
// AI Prompts: Write C# code that uses Aspose.Cells to save a workbook as HTML with a custom data‑cell attribute on each <td> and then checks the file to confirm the attribute exists. | Provide a C# example that uses HtmlAgilityPack instead of regular expressions to verify that every <td> element in the exported HTML contains the data‑cell attribute.

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsHtmlCellAttributeDemo
{
    // Creates a workbook, fills it with sample data, sets HtmlSaveOptions.CellNameAttribute to "data-cell", saves the sheet as HTML, reads the output, and uses pattern matching to confirm that each <td> element contains the required attribute. Reports success or failure.
    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a new workbook and add sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Header1");
                sheet.Cells["B1"].PutValue("Header2");
                sheet.Cells["A2"].PutValue("Row1Col1");
                sheet.Cells["B2"].PutValue("Row1Col2");
                sheet.Cells["A3"].PutValue("Row2Col1");
                sheet.Cells["B3"].PutValue("Row2Col2");

                // 2. Configure HTML save options to write a custom attribute for each cell
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    CellNameAttribute = "data-cell",          // e.g., <td data-cell='A1'>...</td>
                    ExportActiveWorksheetOnly = true          // Export only the active sheet
                };

                // 3. Save the workbook as HTML
                string htmlPath = "output.html";
                workbook.Save(htmlPath, saveOptions);

                // Ensure the HTML file was created
                if (!File.Exists(htmlPath))
                {
                    Console.WriteLine($"Failed to create HTML file at '{htmlPath}'.");
                    return;
                }

                // 4. Load the generated HTML content
                string htmlContent = File.ReadAllText(htmlPath);

                // 5. Find all <td> elements using a simple regex
                MatchCollection tdMatches = Regex.Matches(htmlContent, @"<td\b[^>]*>", RegexOptions.IgnoreCase);
                if (tdMatches.Count == 0)
                {
                    Console.WriteLine("No <td> elements found in the HTML.");
                    return;
                }

                // 6. Verify each <td> has the required attribute
                bool allHaveAttribute = true;
                foreach (Match match in tdMatches)
                {
                    string tdTag = match.Value;
                    if (!Regex.IsMatch(tdTag, @"\bdata-cell\s*=", RegexOptions.IgnoreCase))
                    {
                        allHaveAttribute = false;
                        Console.WriteLine($"Missing attribute in cell HTML: {tdTag}");
                    }
                }

                // 7. Output verification result
                if (allHaveAttribute)
                    Console.WriteLine("Verification succeeded: every <td> element contains the 'data-cell' attribute.");
                else
                    Console.WriteLine("Verification failed: some <td> elements are missing the 'data-cell' attribute.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
