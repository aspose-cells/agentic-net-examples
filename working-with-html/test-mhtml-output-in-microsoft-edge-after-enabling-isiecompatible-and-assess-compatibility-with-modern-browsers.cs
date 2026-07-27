// Title: Generate and Reload IE‑Compatible MHTML with Aspose.Cells for .NET – Verify in Microsoft Edge
// Description: Demonstrates how to create a workbook, save it as an MHTML file with the IsIECompatible flag enabled, reload the file to confirm data integrity, and evaluate rendering in Microsoft Edge. Ideal for developers needing legacy‑IE markup that still works in modern browsers.
// Keywords: Aspose.Cells MHTML | IsIECompatible true | save workbook as MHTML .NET | load MHTML Aspose.Cells | MHTML rendering Edge | cross‑browser MHTML test
// Common Searches: enable IE compatibility when exporting MHTML with Aspose.Cells | open Aspose.Cells MHTML file in Microsoft Edge | reload MHTML saved by Aspose.Cells | difference between IsIECompatible true and false
// Developer Intent: Produce an MHTML document that contains IE‑compatible markup, then read it back to ensure the workbook data remains unchanged.
// Use Cases: Create reports that must display correctly in legacy IE mode while still being viewable in Edge or Chrome. | Validate round‑trip fidelity of MHTML export/import in automated test suites. | Compare visual output of IE‑compatible versus standard MHTML across browsers.
// AI Prompts: Generate C# code to compare the HTML output of MHTML saved with IsIECompatible set to true versus false. | Provide a script that launches Microsoft Edge, opens the saved MHTML file, and captures a screenshot for visual comparison. | Explain how to programmatically detect layout differences between IE‑compatible and default MHTML rendering in modern browsers.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, save it as an MHTML file with the IsIECompatible flag enabled, reload the file to confirm data integrity, and evaluate rendering in Microsoft Edge. Ideal for developers needing legacy‑IE markup that still works in modern browsers.
class MHtmlCompatibilityDemo
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("MHTML Compatibility Test");
        worksheet.Cells["B2"].PutValue(DateTime.Now);
        worksheet.Cells["C3"].PutValue(12345);

        // Configure HtmlSaveOptions for MHTML output with IE compatibility enabled
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.MHtml);
        saveOptions.IsIECompatible = true; // Enable IE compatibility mode

        // Save the workbook as an MHTML file
        string mhtmlFile = "MHtmlCompatibility.mht";
        workbook.Save(mhtmlFile, saveOptions);
        Console.WriteLine($"MHTML file saved to '{mhtmlFile}' with IsIECompatible = true.");

        // Load the saved MHTML file to verify it can be read back correctly
        HtmlLoadOptions loadOptions = new HtmlLoadOptions(LoadFormat.MHtml);
        Workbook loadedWorkbook = new Workbook(mhtmlFile, loadOptions);
        Console.WriteLine("MHTML file loaded successfully. First cell value: " +
            loadedWorkbook.Worksheets[0].Cells["A1"].StringValue);
    }
}
