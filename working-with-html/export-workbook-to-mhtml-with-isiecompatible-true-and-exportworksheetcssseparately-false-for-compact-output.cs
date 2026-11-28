// Title: Export an Aspose.Cells workbook to a single‑file MHTML with IE compatibility and embedded CSS using C#
// Common Searches: c# aspose.cells save workbook as mhtml with IsIECompatible true | how to generate compact MHTML from Excel using Aspose.Cells HtmlSaveOptions | asp.net export Excel to .mht with embedded images base64 | disable separate CSS files when exporting workbook to MHTML with Aspose.Cells | create IE‑compatible MHTML output from a workbook in C#
// Tags: Aspose.Cells MHTML export with IE compatibility | HtmlSaveOptions ExportWorksheetCSSSeparately false | Export images as Base64 in MHTML Aspose.Cells | compact single‑file MHTML output C# | save workbook as .mht using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace ExportWorkbookToMhtml
{
    // The example creates a workbook, fills it with sample data, configures HtmlSaveOptions for MHTML with IsIECompatible enabled, disables separate CSS files, embeds images as Base64, and saves the result as a self‑contained .mht file while ensuring the output directory exists and handling exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and add some sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "SampleData";

                // Populate the worksheet with sample values
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Quantity");
                sheet.Cells["C1"].PutValue("Price");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["C2"].PutValue(0.5);
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["B3"].PutValue(85);
                sheet.Cells["C3"].PutValue(0.3);
                sheet.Cells["A4"].PutValue("Cherry");
                sheet.Cells["B4"].PutValue(60);
                sheet.Cells["C4"].PutValue(1.2);

                // Configure HTML save options for MHTML output
                HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.MHtml)
                {
                    // Enable IE compatibility mode
                    IsIECompatible = true,

                    // Export CSS together with the HTML (compact output)
                    ExportWorksheetCSSSeparately = false,

                    // Embed images as Base64 to keep the MHTML self‑contained
                    ExportImagesAsBase64 = true
                };

                // Define output file path
                string outputPath = "WorkbookExport.mht";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as MHTML using the configured options
                workbook.Save(outputPath, saveOptions);

                Console.WriteLine($"Workbook successfully exported to MHTML at: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("An error occurred during export:");
                Console.Error.WriteLine(ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
