// Title: Enable IE‑compatible MHTML export with worksheet tabs using Aspose.Cells HtmlSaveOptions in C#
// Common Searches: Aspose.Cells C# enable worksheet tab switching in MHTML for Internet Explorer | Set HtmlSaveOptions.IsIECompatible to true when saving workbook as MHTML | Export Excel workbook to MHTML with tabs visible in IE using Aspose.Cells | How to embed images as base64 in MHTML with Aspose.Cells | Save all worksheets to a single MHTML file in C# with Aspose.Cells
// Tags: HtmlSaveOptions.IsIECompatible property | MHTML export preserving worksheet navigation | Aspose.Cells export to MHTML C# | Embed images as base64 in MHTML | Export all worksheets to MHTML | Internet Explorer support for MHTML

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMhtmlExample
{
    // The example loads or creates an Excel workbook, adds a second worksheet, configures HtmlSaveOptions for MHTML with IsIECompatible set to true, disables ExportActiveWorksheetOnly, embeds images as base64, ensures the output folder exists, and saves the workbook as an MHTML file that retains worksheet tabs and works in Internet Explorer.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for input and output files
            string inputPath = "input.xlsx";
            string outputPath = "output.mht";

            try
            {
                Workbook workbook;

                // Load existing workbook if it exists; otherwise create a new one with sample data
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                    Console.WriteLine($"Loaded workbook from \"{inputPath}\".");
                }
                else
                {
                    workbook = new Workbook();
                    Worksheet sheet = workbook.Worksheets[0];
                    sheet.Name = "SampleData";

                    // Populate some sample data
                    sheet.Cells["A1"].PutValue("Product");
                    sheet.Cells["B1"].PutValue("Price");
                    sheet.Cells["A2"].PutValue("Apple");
                    sheet.Cells["B2"].PutValue(1.20);
                    sheet.Cells["A3"].PutValue("Banana");
                    sheet.Cells["B3"].PutValue(0.80);
                    sheet.Cells["A4"].PutValue("Cherry");
                    sheet.Cells["B4"].PutValue(2.50);

                    // Add a second worksheet to demonstrate tab switching
                    Worksheet sheet2 = workbook.Worksheets.Add("Summary");
                    sheet2.Cells["A1"].PutValue("Total");
                    sheet2.Cells["B1"].Formula = "=SUM(SampleData!B2:B4)";

                    Console.WriteLine("Created a new workbook with sample data.");
                }

                // Configure HTML save options for MHTML output
                HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.MHtml)
                {
                    // Enable IE compatibility so that worksheet tabs can be switched in Internet Explorer
                    IsIECompatible = true,

                    // Export all worksheets (default) – keep tabs visible in the MHTML file
                    ExportActiveWorksheetOnly = false,

                    // Embed images as base64 to keep the MHTML self‑contained
                    ExportImagesAsBase64 = true
                };

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as MHTML using the configured options
                workbook.Save(outputPath, saveOptions);
                Console.WriteLine($"Workbook successfully saved as MHTML to \"{outputPath}\" with IE compatibility enabled.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during processing:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
