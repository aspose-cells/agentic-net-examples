// Title: Save an Aspose.Cells Workbook to HTML with Formulas and Reload It to Verify Round‑Trip Fidelity (C#)
// Description: Demonstrates how to export a workbook’s active worksheet to HTML while preserving formulas using HtmlSaveOptions, then reload the HTML with HtmlLoadOptions.LoadFormulas and compare original and loaded cell values to confirm data integrity.
// Keywords: Aspose.Cells C# HTML export | HtmlSaveOptions ExportFormula | HtmlLoadOptions LoadFormulas | Excel to HTML round trip | verify workbook fidelity | load HTML into workbook Aspose | preserve formulas in HTML | export active worksheet only | Aspose.Cells roundtrip test | C# Aspose.Cells example
// Common Searches: Aspose.Cells export worksheet to HTML with formulas | How to load HTML back into a workbook preserving formulas | Round‑trip Excel to HTML and back using Aspose.Cells | C# HtmlSaveOptions ExportFormula example | HtmlLoadOptions LoadFormulas usage
// Developer Intent: The developer needs to save a workbook as HTML, keep formulas intact during the export, reload the HTML into a new workbook, and verify that the original and reloaded content match.
// Use Cases: Create an HTML preview of a sheet for web display while retaining editable formulas for later processing. | Automate a regression test that checks for data or formula loss after converting Excel to HTML and back. | Export a single active worksheet for reporting, then re‑import it into a fresh workbook for further calculations.
// AI Prompts: Generate C# code that uses Aspose.Cells to save a workbook to HTML with formulas and then reload the HTML preserving those formulas. | Explain the impact of HtmlLoadOptions.LoadFormulas on importing HTML into a Workbook and note any limitations. | Provide a method to compare cell values and formulas after an HTML round‑trip to confirm fidelity.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to export a workbook’s active worksheet to HTML while preserving formulas using HtmlSaveOptions, then reload the HTML with HtmlLoadOptions.LoadFormulas and compare original and loaded cell values to confirm data integrity.
class Program
{
    static void Main()
    {
        try
        {
            // Create a workbook and add sample data
            Workbook original = new Workbook();
            Worksheet ws = original.Worksheets[0];
            ws.Cells["A1"].PutValue("Hello");
            ws.Cells["B1"].PutValue(123);
            ws.Cells["C1"].Formula = "=B1+10";

            // Configure HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                ExportActiveWorksheetOnly = true, // export only the active sheet
                ExportFormula = true               // keep formulas in the HTML
            };

            // Save the workbook as HTML
            string htmlPath = "roundtrip.html";
            original.Save(htmlPath, saveOptions);

            // Ensure the HTML file exists before loading
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Error: File '{htmlPath}' not found.");
                return;
            }

            // Configure HTML load options
            HtmlLoadOptions loadOptions = new HtmlLoadOptions
            {
                LoadFormulas = true // import formulas from HTML
            };

            // Load the HTML back into a new workbook
            Workbook loaded = new Workbook(htmlPath, loadOptions);
            Worksheet loadedWs = loaded.Worksheets[0];

            // Verify round‑trip fidelity by comparing cell values and formulas
            Console.WriteLine("Original A1: " + ws.Cells["A1"].StringValue);
            Console.WriteLine("Loaded   A1: " + loadedWs.Cells["A1"].StringValue);

            // B1 may be empty after load; handle safely
            Cell originalB1 = ws.Cells["B1"];
            Cell loadedB1 = loadedWs.Cells["B1"];
            Console.WriteLine("Original B1: " + originalB1.IntValue);
            Console.WriteLine("Loaded   B1: " + (loadedB1.Type == CellValueType.IsNull ? "Empty" : loadedB1.IntValue.ToString()));

            Console.WriteLine("Original C1 formula: " + ws.Cells["C1"].Formula);
            Console.WriteLine("Loaded   C1 formula: " + loadedWs.Cells["C1"].Formula);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
