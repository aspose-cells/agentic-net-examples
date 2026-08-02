// Title: Create IE‑11 Compatible MHTML with Worksheet Tabs using Aspose.Cells for .NET
// Description: C# example that builds a workbook with two sheets, sets ShowTabs, enables HtmlSaveOptions.IsIECompatible, and saves as .mht so the tabs work in Internet Explorer 11.
// Keywords: Aspose.Cells | C# MHTML export | Internet Explorer 11 | IsIECompatible | worksheet tabs | HtmlSaveOptions | .NET Excel to MHTML | save multiple worksheets | IE compatibility mode | MHT file generation
// Common Searches: Aspose.Cells save as MHTML IE11 | How to enable worksheet tabs in MHTML export | IsIECompatible flag Aspose.Cells | Export Excel to .mht with tabs | C# generate MHTML for Internet Explorer | ShowTabs property Aspose.Cells HTML | Export all worksheets to single MHTML | Test MHTML output in IE11
// Developer Intent: Generate an MHTML file that retains visible worksheet tabs and renders correctly in Internet Explorer 11.
// Use Cases: Legacy reporting portals that need a single MHTML document with clickable sheet tabs. | Automated conversion of Excel workbooks to IE‑compatible web pages. | Testing and validation of Aspose.Cells HTML export settings for tab navigation. | Embedding workbook snapshots in intranet sites accessed via IE11.
// AI Prompts: Write C# code using Aspose.Cells to export a multi‑sheet workbook to a single .mht file with visible tabs for IE 11. | Explain the effect of HtmlSaveOptions.IsIECompatible on the generated MHTML and which additional options must be set for tab navigation. | Provide a step‑by‑step guide to verify that the saved MHTML contains the correct HTML for worksheet tabs when opened in Internet Explorer 11. | Suggest how to programmatically check the presence of the <div class="sheet-tab"> elements in the MHTML output.

using System;
using Aspose.Cells;

namespace AsposeCellsMhtmlIeTest
{
    // C# example that builds a workbook with two sheets, sets ShowTabs, enables HtmlSaveOptions.IsIECompatible, and saves as .mht so the tabs work in Internet Explorer 11.
    class Program
    {
        static void Main()
        {
            // Create a new workbook with two worksheets
            Workbook workbook = new Workbook();
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "FirstSheet";
            sheet1.Cells["A1"].PutValue("Sheet 1 - Cell A1");
            sheet1.Cells["B2"].PutValue("Sheet 1 - Cell B2");

            Worksheet sheet2 = workbook.Worksheets.Add("SecondSheet");
            sheet2.Cells["A1"].PutValue("Sheet 2 - Cell A1");
            sheet2.Cells["C3"].PutValue("Sheet 2 - Cell C3");

            // Ensure workbook tabs are visible (required for tab navigation in the HTML/MHTML output)
            workbook.Settings.ShowTabs = true;

            // Configure HTML save options for MHTML output
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                // Enable IE compatibility mode so the generated MHTML works correctly in Internet Explorer 11
                IsIECompatible = true,

                // Export all worksheets (default), keep tabs visible in the output
                ExportActiveWorksheetOnly = false,

                // Export worksheet properties (default) – required for proper tab navigation
                ExportWorksheetProperties = true
            };

            // Save the workbook as MHTML (the .mht extension triggers MHTML format)
            string outputPath = "WorkbookWithTabs.mht";
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"MHTML file saved to '{outputPath}'. Open it in Internet Explorer 11 to verify tab navigation.");
        }
    }
}
