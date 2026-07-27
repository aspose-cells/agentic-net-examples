// Title: Aspose.Cells .NET: Change TableCssId after HTML export and verify the update
// Description: Demonstrates how to export a workbook to HTML with an initial TableCssId, modify the TableCssId on the same HtmlSaveOptions instance, re‑export to a second file, and programmatically confirm that the new TableCssId appears in the generated HTML.
// Keywords: Aspose.Cells | C# | .NET | HtmlSaveOptions | TableCssId | HTML export | dynamic CSS class | re‑export workbook | verify HTML content | programmatic validation
// Common Searches: Aspose.Cells change TableCssId after export | HTML export with custom table CSS id .NET | verify TableCssId in exported HTML | re‑save workbook with new TableCssId | C# Aspose.Cells HTML styling
// Developer Intent: Update the TableCssId used for HTML export of a workbook and ensure the change is reflected in the output file.
// Use Cases: Apply a specific CSS class to the HTML table generated from a workbook. | Switch table styling on‑the‑fly by changing TableCssId and re‑exporting without recreating the workbook. | Automated testing that confirms the exported HTML contains the expected TableCssId value.
// AI Prompts: Write C# code that exports an Aspose.Cells workbook to HTML with a given TableCssId, then changes the TableCssId and creates a second HTML file. | Provide a method to read an exported HTML file and assert that it includes the updated TableCssId string. | Explain how TableCssId influences the HTML output and how to use it for dynamic styling in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTableCssIdDemo
{
    // Demonstrates how to export a workbook to HTML with an initial TableCssId, modify the TableCssId on the same HtmlSaveOptions instance, re‑export to a second file, and programmatically confirm that the new TableCssId appears in the generated HTML.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["A2"].PutValue("John");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Alice");
            sheet.Cells["B3"].PutValue(25);

            // First export with an initial TableCssId
            HtmlSaveOptions firstOptions = new HtmlSaveOptions(SaveFormat.Html);
            firstOptions.TableCssId = "initial-table-style";
            string firstHtmlPath = "initialExport.html";
            workbook.Save(firstHtmlPath, firstOptions);
            Console.WriteLine($"First HTML saved to '{firstHtmlPath}' with TableCssId = {firstOptions.TableCssId}");

            // Change the TableCssId and re‑export the same workbook
            firstOptions.TableCssId = "updated-table-style";
            string secondHtmlPath = "updatedExport.html";
            workbook.Save(secondHtmlPath, firstOptions);
            Console.WriteLine($"Second HTML saved to '{secondHtmlPath}' with TableCssId = {firstOptions.TableCssId}");

            // Verify that the updated HTML contains the new TableCssId string
            string htmlContent = File.ReadAllText(secondHtmlPath);
            bool containsUpdatedId = htmlContent.Contains(firstOptions.TableCssId);
            Console.WriteLine($"Verification: HTML contains updated TableCssId? {containsUpdatedId}");
        }
    }
}
