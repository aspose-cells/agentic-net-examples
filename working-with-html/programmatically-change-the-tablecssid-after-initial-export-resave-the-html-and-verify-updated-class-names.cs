// Title: Update TableCssId After HTML Export with Aspose.Cells (C#) – Re‑save and Verify
// Description: Demonstrates how to export a workbook to HTML using Aspose.Cells, change the HtmlSaveOptions.TableCssId, re‑export the file, and confirm the new CSS identifier appears. The example also shows loading the generated HTML back into a workbook to illustrate round‑trip behavior.
// Keywords: Aspose.Cells HtmlSaveOptions TableCssId | C# export workbook to HTML | modify TableCssId after export | verify HTML CSS id Aspose | HtmlLoadOptions round‑trip | re‑save HTML with custom CSS id
// Common Searches: how to set TableCssId in Aspose.Cells HTML export | change CSS id after saving workbook to HTML | verify custom TableCssId in exported HTML | load exported HTML back into Aspose.Cells workbook | C# Aspose.Cells re‑save HTML with new TableCssId
// Developer Intent: Set a custom TableCssId for an HTML export, re‑save the workbook, and ensure the updated CSS identifier is present in the generated HTML.
// Use Cases: Export a workbook to HTML with default settings and confirm no TableCssId is added. | Assign a custom value to HtmlSaveOptions.TableCssId, re‑export, and validate the CSS id appears in the output. | Load the first HTML file back into a Workbook to demonstrate that TableCssId is not retained during loading.
// AI Prompts: Generate C# code that creates a workbook, saves it to HTML, sets HtmlSaveOptions.TableCssId to a custom string, re‑saves, and checks the HTML for that id. | Write a C# unit test that asserts the TableCssId is missing in the initial HTML export and present after updating the option and re‑exporting. | Explain why HtmlLoadOptions does not preserve TableCssId when loading an HTML file back into a Workbook and how to handle this scenario.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to export a workbook to HTML using Aspose.Cells, change the HtmlSaveOptions.TableCssId, re‑export the file, and confirm the new CSS identifier appears. The example also shows loading the generated HTML back into a workbook to illustrate round‑trip behavior.
class Program
{
    static void Main()
    {
        // 1. Create a workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Age");
        sheet.Cells["A2"].PutValue("John");
        sheet.Cells["B2"].PutValue(30);

        // 2. Export the workbook to HTML with the default TableCssId (empty)
        string firstHtmlPath = "first.html";
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
        workbook.Save(firstHtmlPath, saveOptions);
        Console.WriteLine($"First HTML saved to: {firstHtmlPath}");

        // 3. Load the generated HTML back into a workbook (demonstrates the load step)
        HtmlLoadOptions loadOptions = new HtmlLoadOptions();
        Workbook loadedWorkbook = new Workbook(firstHtmlPath, loadOptions);
        Console.WriteLine($"Loaded HTML workbook contains {loadedWorkbook.Worksheets.Count} worksheet(s).");

        // 4. Verify that the default TableCssId does NOT appear in the first HTML
        string firstHtmlContent = File.ReadAllText(firstHtmlPath);
        bool firstContainsCustomId = firstHtmlContent.Contains("custom-table-style");
        Console.WriteLine($"First HTML contains custom TableCssId? {firstContainsCustomId}");

        // 5. Change TableCssId and re‑save the original workbook
        saveOptions.TableCssId = "custom-table-style";
        string secondHtmlPath = "second.html";
        workbook.Save(secondHtmlPath, saveOptions);
        Console.WriteLine($"Second HTML saved to: {secondHtmlPath}");

        // 6. Verify that the updated TableCssId appears in the second HTML
        string secondHtmlContent = File.ReadAllText(secondHtmlPath);
        bool secondContainsCustomId = secondHtmlContent.Contains("custom-table-style");
        Console.WriteLine($"Second HTML contains updated TableCssId? {secondContainsCustomId}");
    }
}
