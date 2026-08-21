// Title: C# Aspose.Cells: Export Excel to HTML with custom data-cell attribute and verify each <td>
// Description: Shows how to build a workbook, configure HtmlSaveOptions.CellNameAttribute to add a data-cell attribute (e.g., data-cell="A1") to every <td> in the exported HTML, save the file, and programmatically confirm that the attribute exists on all table cells.
// Keywords: Aspose.Cells C# | HtmlSaveOptions CellNameAttribute | data-cell attribute | export Excel to HTML | HTML table cell verification | C# HTML parsing | HtmlAgilityPack example | .NET Excel to HTML conversion | custom HTML attributes Aspose | automated HTML validation
// Common Searches: Aspose.Cells add custom attribute to exported HTML cells | C# verify data-cell attribute in HTML generated from Excel | HtmlSaveOptions.CellNameAttribute usage example | How to include Excel cell address in HTML table with Aspose | C# code to check every <td> has a specific attribute
// Developer Intent: Create an HTML file from an Excel workbook where each <td> includes a data-cell attribute, then programmatically ensure the attribute is present on every cell.
// Use Cases: Link client‑side scripts to original Excel coordinates for interactive dashboards. | Implement regression tests that detect missing custom attributes after library upgrades. | Produce searchable HTML tables where each cell can be uniquely identified by its source address.
// AI Prompts: Generate C# code using Aspose.Cells to export a worksheet to HTML with a data-cell attribute on each <td> and then validate the output. | Write a unit test in C# that loads the exported HTML and asserts that every table cell contains a data-cell attribute matching its Excel reference. | Suggest a robust HTML parsing method (e.g., HtmlAgilityPack) to scan the Aspose.Cells HTML output for missing data-cell attributes and report discrepancies.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlCellAttributeDemo
{
    // Shows how to build a workbook, configure HtmlSaveOptions.CellNameAttribute to add a data-cell attribute (e.g., data-cell="A1") to every <td> in the exported HTML, save the file, and programmatically confirm that the attribute exists on all table cells.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and add sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Header1");
                sheet.Cells["B1"].PutValue("Header2");
                sheet.Cells["A2"].PutValue("Row1Col1");
                sheet.Cells["B2"].PutValue("Row1Col2");
                sheet.Cells["A3"].PutValue("Row2Col1");
                sheet.Cells["B3"].PutValue("Row2Col2");

                // Configure HTML save options to write a custom attribute with the cell name
                HtmlSaveOptions saveOptions = new HtmlSaveOptions();
                saveOptions.CellNameAttribute = "data-cell"; // each <td> will have data-cell="A1", etc.

                // Define output HTML path
                string htmlPath = Path.Combine(Environment.CurrentDirectory, "output.html");

                // Save the workbook as HTML using the configured options
                workbook.Save(htmlPath, saveOptions);

                // Ensure the HTML file was created
                if (!File.Exists(htmlPath))
                {
                    Console.WriteLine("Failed to generate the HTML file.");
                    return;
                }

                // Load the generated HTML content
                string htmlContent = File.ReadAllText(htmlPath);

                // Simple verification: each <td> should contain the data-cell attribute
                bool allHaveAttribute = true;
                string[] tdSegments = htmlContent.Split(new[] { "<td", "</td>" }, StringSplitOptions.None);
                foreach (string segment in tdSegments)
                {
                    int closeIdx = segment.IndexOf('>');
                    if (closeIdx > -1)
                    {
                        string tdTag = segment.Substring(0, closeIdx);
                        if (!tdTag.Contains("data-cell="))
                        {
                            allHaveAttribute = false;
                            Console.WriteLine($"Missing data-cell attribute in cell HTML: <td{tdTag}>...</td>");
                        }
                    }
                }

                Console.WriteLine(allHaveAttribute
                    ? "All table cells contain the data-cell attribute."
                    : "Some table cells are missing the data-cell attribute.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
