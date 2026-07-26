// Title: Enable IE Compatibility for MHTML Export in Aspose.Cells (C#) – Worksheet Tab Switching
// Description: Shows how to set HtmlSaveOptions.IsIECompatible = true before saving a workbook as MHTML, so the resulting file lets users switch worksheet tabs in Internet Explorer.
// Keywords: Aspose.Cells | HtmlSaveOptions | IsIECompatible | MHTML export | C# | worksheet tab switching | Internet Explorer compatibility | Excel to MHTML | legacy browser support | .NET example
// Common Searches: Aspose.Cells enable IE compatibility | HtmlSaveOptions IsIECompatible C# example | MHTML export worksheet tabs Internet Explorer | Save Excel as MHTML with tab navigation | Aspose.Cells MHTML IE support
// Developer Intent: Configure HtmlSaveOptions for MHTML and turn on IE compatibility so the generated file supports worksheet tab navigation in Internet Explorer.
// Use Cases: Export a multi‑sheet Excel workbook to a single MHTML file that retains tab navigation for users of Internet Explorer. | Create legacy‑browser‑friendly reports from Aspose.Cells that can be viewed in older versions of IE. | Provide downloadable MHTML documents with active worksheet tabs for web applications targeting enterprise environments still using IE.
// AI Prompts: Generate a C# snippet that saves an Aspose.Cells workbook as MHTML with worksheet tab switching enabled for Internet Explorer. | Explain the impact of HtmlSaveOptions.IsIECompatible on MHTML output and how to use it in .NET. | Show step‑by‑step how to enable IE compatibility when exporting Excel to MHTML with Aspose.Cells.

using System;
using Aspose.Cells;

// Shows how to set HtmlSaveOptions.IsIECompatible = true before saving a workbook as MHTML, so the resulting file lets users switch worksheet tabs in Internet Explorer.
class EnableIECompatibilityDemo
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook workbook = new Workbook();
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Name = "Sheet1";
        sheet1.Cells["A1"].PutValue("First sheet content");

        // Add a second worksheet to demonstrate tab switching
        int sheet2Index = workbook.Worksheets.Add();
        Worksheet sheet2 = workbook.Worksheets[sheet2Index];
        sheet2.Name = "Sheet2";
        sheet2.Cells["A1"].PutValue("Second sheet content");

        // Create HtmlSaveOptions for MHTML format
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.MHtml);
        // Enable IE compatibility to allow worksheet tab switching in Internet Explorer
        saveOptions.IsIECompatible = true;

        // Save the workbook as MHTML using the configured options
        workbook.Save("output.mht", saveOptions);
    }
}
