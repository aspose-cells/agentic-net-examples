// Title: Preserve DIV Layout When Loading and Saving HTML with Aspose.Cells (C#)
// Description: Demonstrates how to load an HTML file with DIV‑based layout into an Aspose.Cells Workbook using HtmlLoadOptions.SupportDivTag, export it back to HTML with HtmlSaveOptions.ParseHtmlTagInCell disabled, and verify that the DIV structure remains unchanged by comparing tag counts.
// Keywords: Aspose.Cells HTML DIV preservation | HtmlLoadOptions.SupportDivTag C# | HtmlSaveOptions.ParseHtmlTagInCell | load HTML to workbook Aspose | save workbook to HTML without parsing tags | compare DIV count Aspose.Cells | C# round‑trip HTML Excel conversion
// Common Searches: how to keep div tags when converting HTML to Excel with Aspose.Cells | Aspose.Cells preserve div layout example .NET | verify HTML div structure after saving workbook | HtmlLoadOptions SupportDivTag usage | HtmlSaveOptions ParseHtmlTagInCell effect
// Developer Intent: Load an HTML file that uses DIVs, save it back to HTML without altering those DIV elements, and programmatically confirm the layout is unchanged.
// Use Cases: Import a web page that relies on DIV layout into Excel for data processing while retaining its visual structure. | Perform a round‑trip conversion (HTML → Excel → HTML) and ensure the DIV hierarchy is preserved for downstream web rendering. | Create an automated test that validates Aspose.Cells does not modify the number or placement of DIV tags during conversion.
// AI Prompts: Write C# code that loads an HTML file with Aspose.Cells using HtmlLoadOptions.SupportDivTag, saves it preserving DIV tags, and checks the DIV count before and after conversion. | Explain how HtmlSaveOptions.ParseHtmlTagInCell influences DIV preservation when exporting a workbook to HTML. | Generate an MSTest unit test that asserts the original and saved HTML files contain the same number of <div> elements after conversion with Aspose.Cells.

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsDivPreserveDemo
{
    // Demonstrates how to load an HTML file with DIV‑based layout into an Aspose.Cells Workbook using HtmlLoadOptions.SupportDivTag, export it back to HTML with HtmlSaveOptions.ParseHtmlTagInCell disabled, and verify that the DIV structure remains unchanged by comparing tag counts.
    class Program
    {
        static void Main()
        {
            // Paths for the input and output HTML files
            string inputHtmlPath = "input.html";
            string outputHtmlPath = "output.html";

            // Load the HTML file with DIV layout support enabled
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            loadOptions.SupportDivTag = true;
            Workbook workbook = new Workbook(inputHtmlPath, loadOptions);

            // Save the workbook back to HTML while preserving HTML tags in cells
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.ParseHtmlTagInCell = false; // keep <div> tags as they are
            workbook.Save(outputHtmlPath, saveOptions);

            // Verify that the DIV structure remains unchanged
            string originalHtml = File.ReadAllText(inputHtmlPath);
            string savedHtml = File.ReadAllText(outputHtmlPath);

            int originalDivCount = Regex.Matches(originalHtml, "<div\\b", RegexOptions.IgnoreCase).Count;
            int savedDivCount = Regex.Matches(savedHtml, "<div\\b", RegexOptions.IgnoreCase).Count;

            Console.WriteLine($"Original DIV count: {originalDivCount}");
            Console.WriteLine($"Saved DIV count: {savedDivCount}");
            Console.WriteLine("DIV structure preserved: " + (originalDivCount == savedDivCount));
        }
    }
}
