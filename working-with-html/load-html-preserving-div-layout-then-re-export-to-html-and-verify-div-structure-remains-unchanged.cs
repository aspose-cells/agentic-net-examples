// Title: Preserve DIV Layout When Loading and Saving HTML with Aspose.Cells (C#)
// Description: Demonstrates how to enable HtmlLoadOptions.SupportDivTag, load an HTML file into a Workbook, re‑export it to HTML, and verify that the original DIV structure remains unchanged by comparing tag counts.
// Keywords: Aspose.Cells HTML div preservation | HtmlLoadOptions SupportDivTag C# | load HTML with DIV layout Aspose.Cells | save workbook to HTML without losing DIVs | verify DIV count after export | Aspose.Cells HTML to Excel conversion | C# Aspose.Cells HTML example
// Common Searches: keep div elements when converting HTML with Aspose.Cells | Aspose.Cells preserve div layout on save | how to enable SupportDivTag in HtmlLoadOptions | compare div tags before and after Aspose.Cells export | C# load HTML preserving div tags Aspose
// Developer Intent: Ensure that the DIV‑based layout of an HTML document is retained after loading it into a Workbook and saving it back with Aspose.Cells.
// Use Cases: Import a web‑based report that uses DIVs, modify data programmatically, and export it without breaking the original layout. | Validate that a spreadsheet template delivered as HTML keeps its DIV structure after processing. | Automate a workflow that reads HTML invoices, updates cells, and writes the file back while preserving the original markup.
// AI Prompts: Write C# code that loads an HTML file with Aspose.Cells, enables SupportDivTag, updates a cell value, and saves the workbook while keeping the original DIV markup. | Create a C# unit test that asserts the number of <div> tags is identical before and after saving a workbook using HtmlLoadOptions.SupportDivTag. | Explain the effect of HtmlLoadOptions.SupportDivTag on HTML‑to‑Excel conversion and list any known limitations.

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// Demonstrates how to enable HtmlLoadOptions.SupportDivTag, load an HTML file into a Workbook, re‑export it to HTML, and verify that the original DIV structure remains unchanged by comparing tag counts.
class PreserveDivLayoutExample
{
    static void Main()
    {
        // Paths to the input and output HTML files
        string inputHtmlPath = "input.html";
        string outputHtmlPath = "output.html";

        // ---------- Load HTML with DIV layout support ----------
        // Create HtmlLoadOptions and enable SupportDivTag
        HtmlLoadOptions loadOptions = new HtmlLoadOptions();
        loadOptions.SupportDivTag = true;

        // Load the HTML file into a Workbook using the load options
        Workbook workbook = new Workbook(inputHtmlPath, loadOptions);

        // ---------- Save the workbook back to HTML ----------
        // Create HtmlSaveOptions (default constructor is sufficient)
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);

        // Save the workbook as HTML
        workbook.Save(outputHtmlPath, saveOptions);

        // ---------- Verify that DIV structure is unchanged ----------
        // Read both original and saved HTML content
        string originalHtml = File.ReadAllText(inputHtmlPath);
        string savedHtml = File.ReadAllText(outputHtmlPath);

        // Count the number of <div> tags in each file (case‑insensitive)
        int originalDivCount = Regex.Matches(originalHtml, "<div\\b", RegexOptions.IgnoreCase).Count;
        int savedDivCount = Regex.Matches(savedHtml, "<div\\b", RegexOptions.IgnoreCase).Count;

        // Output verification result
        Console.WriteLine($"Original <div> count: {originalDivCount}");
        Console.WriteLine($"Saved <div> count: {savedDivCount}");
        Console.WriteLine("DIV structure preserved: " + (originalDivCount == savedDivCount));
    }
}
