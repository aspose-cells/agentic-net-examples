// Title: Convert Aspose.Cells Workbook to HTML and Reload with HtmlLoadOptions (C#)
// Description: Demonstrates how to create a workbook, save it as a self‑contained HTML file using HtmlSaveOptions (preserving properties, formulas, and embedding images), then reload the HTML with HtmlLoadOptions (LoadFormulas enabled) and compare original and loaded cell values to confirm round‑trip fidelity in .NET.
// Keywords: Aspose.Cells HTML export | HtmlSaveOptions C# | HtmlLoadOptions load workbook | Excel to HTML round trip | preserve formulas Aspose.Cells | embed images base64 Aspose | verify data integrity HTML | C# Aspose.Cells example
// Common Searches: save Aspose.Cells workbook as HTML and reload | load HTML workbook with formulas using Aspose.Cells | round‑trip Excel to HTML verification .NET | self‑contained HTML export Aspose.Cells | compare original and loaded workbook cells
// Developer Intent: Export a workbook to HTML, then load it back to ensure that all cell values, dates, numbers, and formulas remain unchanged.
// Use Cases: Generate a single HTML report with embedded images for email or web publishing while keeping formulas for later editing. | Automated testing of Excel‑to‑HTML conversion to guarantee data integrity across export/import cycles. | Create a portable HTML snapshot of a workbook that can be re‑imported into Excel without loss of calculations.
// AI Prompts: Write C# code that saves an Aspose.Cells workbook to a self‑contained HTML file with embedded images and then reloads it preserving formulas using HtmlLoadOptions. | Explain how to compare cell values and formulas after loading an HTML workbook to verify round‑trip fidelity. | Provide troubleshooting steps when formulas disappear after loading an HTML file with HtmlLoadOptions in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlRoundTrip
{
    // Demonstrates how to create a workbook, save it as a self‑contained HTML file using HtmlSaveOptions (preserving properties, formulas, and embedding images), then reload the HTML with HtmlLoadOptions (LoadFormulas enabled) and compare original and loaded cell values to confirm round‑trip fidelity in .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Step 1: Create a new workbook and add sample data.
                Workbook originalWorkbook = new Workbook();
                Worksheet sheet = originalWorkbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Round‑Trip Test");
                sheet.Cells["B2"].PutValue(12345);
                sheet.Cells["C3"].PutValue(DateTime.Now);
                sheet.Cells["D4"].Formula = "=B2*2";

                // Step 2: Save the workbook as HTML using HtmlSaveOptions.
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    ExportWorkbookProperties = true, // Preserve workbook properties.
                    ExportFormula = true,            // Keep formulas in the HTML.
                    ExportActiveWorksheetOnly = true, // Avoid frames; embed the sheet directly.
                    ExportImagesAsBase64 = true      // Embed images to keep HTML self‑contained.
                };

                string htmlFilePath = "RoundTrip.html";
                originalWorkbook.Save(htmlFilePath, saveOptions);

                // Ensure the HTML file was created before attempting to load it.
                if (!File.Exists(htmlFilePath))
                {
                    Console.WriteLine($"Error: The file '{htmlFilePath}' was not found.");
                    return;
                }

                // Step 3: Load the previously saved HTML file using HtmlLoadOptions.
                HtmlLoadOptions loadOptions = new HtmlLoadOptions
                {
                    LoadFormulas = true // Reload formulas.
                    // LoadData is not required; data is loaded by default.
                };

                Workbook loadedWorkbook;
                try
                {
                    loadedWorkbook = new Workbook(htmlFilePath, loadOptions);
                }
                catch (Exception loadEx)
                {
                    Console.WriteLine($"Failed to load HTML workbook: {loadEx.Message}");
                    return;
                }

                // Step 4: Verify round‑trip fidelity by comparing cell values and formulas.
                Worksheet loadedSheet = loadedWorkbook.Worksheets[0];

                Console.WriteLine("Verification Results:");
                Console.WriteLine($"A1 Value  : Original = {sheet.Cells["A1"].StringValue}, Loaded = {loadedSheet.Cells["A1"].StringValue}");
                Console.WriteLine($"B2 Value  : Original = {sheet.Cells["B2"].IntValue}, Loaded = {loadedSheet.Cells["B2"].IntValue}");
                Console.WriteLine($"C3 Value  : Original = {sheet.Cells["C3"].DateTimeValue}, Loaded = {loadedSheet.Cells["C3"].DateTimeValue}");
                Console.WriteLine($"D4 Formula: Original = {sheet.Cells["D4"].Formula}, Loaded = {loadedSheet.Cells["D4"].Formula}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
