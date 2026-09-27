// Title: Create an MHTML file from an Aspose.Cells workbook in C# and test rendering in Google Chrome with IE compatibility disabled
// AI Prompts: Write a C# console application that builds a workbook, populates cells, and saves it as an .mhtml file using HtmlSaveOptions with IsIECompatible set to false and images embedded as Base64. | Configure HtmlSaveOptions to export only the active worksheet, omit hidden worksheets, and disable IE‑specific features, then launch Google Chrome from the code to open the generated MHTML file. | Extend the program to capture a screenshot of Chrome's display of the MHTML file to automatically verify that standard HTML rendering is achieved.
// Common Searches: Aspose.Cells C# export workbook to MHTML without IE compatibility mode | How to set IsIECompatible false in HtmlSaveOptions for Chrome rendering | Embedding images as Base64 in MHTML using Aspose.Cells .NET | Open generated .mhtml file in Google Chrome from C# code | Verify standard HTML rendering of Aspose.Cells MHTML output in Chrome
// Tags: Aspose.Cells HtmlSaveOptions IsIECompatible false | C# export workbook to MHTML | MHTML Base64 image embedding Aspose.Cells | Chrome rendering verification Aspose.Cells MHTML | export active worksheet only Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using System.Drawing.Imaging;

// The example creates a workbook, adds sample data, and saves it as an MHTML file with images embedded as Base64 using HtmlSaveOptions configured to disable IE compatibility, export only the active sheet, and skip hidden sheets. The resulting .mhtml can be opened in Google Chrome to confirm standard rendering without IE‑specific behavior.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and give it a name
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sample";

            // Add some sample data
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue("Row 1");
            sheet.Cells["A3"].PutValue("Row 2");

            // Define output path
            string outputPath = "SampleOutput.mhtml";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as MHTML (HTML with embedded resources)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportImagesAsBase64 = true,          // Embed images directly
                ExportActiveWorksheetOnly = true,    // Export only the active sheet
                ExportHiddenWorksheet = false        // Do not export hidden worksheets
                // ExportChartImageFormat removed for compatibility with current API
            };

            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"MHTML file saved to {outputPath}. Open it in Chrome to verify standard rendering.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
