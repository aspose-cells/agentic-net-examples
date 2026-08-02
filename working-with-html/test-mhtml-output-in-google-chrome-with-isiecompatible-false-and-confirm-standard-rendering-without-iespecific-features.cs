// Title: Aspose.Cells C# – Save Workbook as MHTML without IE Compatibility and Verify Chrome Rendering
// Description: This example creates a workbook, adds sample data, and uses HtmlSaveOptions to export it as MHTML with IsIECompatible set to false and HtmlVersion set to HTML5. The file is then reloaded with HtmlLoadOptions to confirm validity, and instructions are provided to open the .mht file in Google Chrome to ensure standard, non‑IE rendering.
// Keywords: Aspose.Cells MHTML export | IsIECompatible false | HtmlVersion Html5 | Chrome MHTML rendering | C# Aspose.Cells HtmlSaveOptions | load MHTML Aspose.Cells | disable IE compatibility Aspose
// Common Searches: How to export Excel to MHTML with Aspose.Cells C# | Save workbook as MHTML without IE compatibility | Test MHTML output in Google Chrome using Aspose.Cells | Aspose.Cells HtmlSaveOptions IsIECompatible example | Reload MHTML file with Aspose.Cells
// Developer Intent: Generate an MHTML file from a workbook with IE‑specific features disabled and confirm it renders correctly in modern browsers such as Chrome.
// Use Cases: Produce MHTML reports that follow HTML5 standards for consistent display across browsers. | Validate that MHTML files created by Aspose.Cells can be re‑opened without errors. | Automate testing to ensure the exported MHTML contains no IE conditional comments.
// AI Prompts: Give C# code that saves an Aspose.Cells workbook as MHTML with IsIECompatible set to false and opens the file in Chrome for visual verification. | Show how to programmatically scan the generated MHTML for IE conditional comments using Aspose.Cells. | Explain how to compare Chrome and IE rendering of an MHTML file with automated tools.

using System;
using Aspose.Cells;

namespace AsposeCellsMhtmlTest
{
    // This example creates a workbook, adds sample data, and uses HtmlSaveOptions to export it as MHTML with IsIECompatible set to false and HtmlVersion set to HTML5. The file is then reloaded with HtmlLoadOptions to confirm validity, and instructions are provided to open the .mht file in Google Chrome to ensure standard, non‑IE rendering.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("MHTML Test");
            sheet.Cells["A2"].PutValue(DateTime.Now);
            sheet.Cells["B1"].PutValue("Chrome rendering without IE compatibility");

            // Configure HTML save options for MHTML output
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.MHtml);
            // Disable IE compatibility mode to produce standard HTML
            saveOptions.IsIECompatible = false;
            // Use HTML5 standard (optional, improves compatibility)
            saveOptions.HtmlVersion = HtmlVersion.Html5;

            // Save the workbook as an MHTML file
            string outputPath = "MhtmlTestOutput.mht";
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"MHTML file saved to {outputPath} with IsIECompatible = false.");

            // Load the generated MHTML file to verify it can be read back
            HtmlLoadOptions loadOptions = new HtmlLoadOptions(LoadFormat.MHtml);
            Workbook loadedWorkbook = new Workbook(outputPath, loadOptions);
            Console.WriteLine("MHTML file loaded successfully, confirming it is a valid format.");

            // To complete the test, open the saved .mht file in Google Chrome.
            // The page should render using standard HTML features without IE‑specific conditional comments.
        }
    }
}
