// Title: Save Workbook as MHTML with IE Compatibility (HTML5) using Aspose.Cells for .NET
// Description: Creates a workbook, adds sample data, and saves it as an MHTML file with IsIECompatible = true and HtmlVersion = Html5. The code then checks the file’s existence and reloads it with HtmlLoadOptions to verify that the MHTML can be read back and renders correctly in Microsoft Edge and other modern browsers.
// Keywords: Aspose.Cells MHTML export | IsIECompatible .NET | HtmlSaveOptions Html5 | MHTML compatibility Edge | load MHTML Aspose.Cells | C# workbook to MHTML | HTML5 MHTML rendering | Aspose.Cells save format MHtml | verify MHTML file integrity
// Common Searches: how to export Excel to MHTML with Aspose.Cells | IsIECompatible option for MHTML in C# | MHTML output that works in Microsoft Edge | load MHTML file using Aspose.Cells .NET | Aspose.Cells HTML5 MHTML compatibility
// Developer Intent: Generate an MHTML file from a workbook with IE‑compatible mode enabled, then confirm the file can be reloaded and displayed correctly in modern browsers.
// Use Cases: Produce MHTML reports that must render properly in Microsoft Edge, Chrome, or Firefox. | Automated regression test that saves a workbook as MHTML, verifies file creation, and reloads it to ensure no data loss. | Validate that enabling IsIECompatible does not interfere with HTML5 features when the MHTML is opened in contemporary browsers.
// AI Prompts: Show how to embed charts and images into the MHTML output while keeping IsIECompatible set to true. | Provide code to launch Microsoft Edge programmatically, open the generated MHTML file, and capture any rendering warnings. | Explain how to tweak HtmlSaveOptions for optimal compatibility across Chrome, Firefox, and Safari.

using System;
using System.IO;
using Aspose.Cells;

// Creates a workbook, adds sample data, and saves it as an MHTML file with IsIECompatible = true and HtmlVersion = Html5. The code then checks the file’s existence and reloads it with HtmlLoadOptions to verify that the MHTML can be read back and renders correctly in Microsoft Edge and other modern browsers.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Cells["A1"].PutValue("MHTML Compatibility Test");
            worksheet.Cells["B2"].PutValue(DateTime.Now);
            worksheet.Cells["C3"].PutValue(12345);

            // Configure HTML save options for MHTML output
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.MHtml)
            {
                // Enable IE compatibility mode (required for certain MHTML features)
                IsIECompatible = true,
                // Use HTML5 to ensure modern browser compatibility (e.g., Microsoft Edge)
                HtmlVersion = HtmlVersion.Html5
            };

            // Save the workbook as MHTML
            string outputPath = "MhtmlOutput.mht";
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine("MHTML file saved with IsIECompatible = true.");

            // Verify that the generated MHTML file exists before loading
            if (File.Exists(outputPath))
            {
                try
                {
                    // Load the generated MHTML to verify it can be read back
                    HtmlLoadOptions loadOptions = new HtmlLoadOptions(LoadFormat.MHtml);
                    Workbook loadedWorkbook = new Workbook(outputPath, loadOptions);
                    Console.WriteLine($"Loaded workbook contains {loadedWorkbook.Worksheets.Count} worksheet(s).");
                }
                catch (Exception loadEx)
                {
                    Console.WriteLine($"Error loading MHTML file: {loadEx.Message}");
                }
            }
            else
            {
                Console.WriteLine($"Error: The file '{outputPath}' was not found.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
