// Title: Generate IE‑compatible MHTML from a multi‑sheet workbook and confirm worksheet tabs using Aspose.Cells for .NET
// Common Searches: c# export workbook with multiple sheets to .mht using Aspose.Cells | how to include worksheet tabs when saving Excel as MHTML for Internet Explorer | Aspose.Cells HtmlSaveOptions configuration for IE compatible MHTML output | verify sheet tab markup in generated MHTML file with C# regex | save Excel as MHTML with base64 images and grid lines using Aspose.Cells | automatically open generated .mht file after export in a .NET application
// Tags: Aspose.Cells export all worksheets to MHTML | IE compatible MHTML output settings | HtmlSaveOptions grid lines UTF‑8 encoding | embed images as Base64 in MHTML | sheet tab markup verification in MHTML | C# launch .mht file with default viewer

using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsMhtmlDemo
{
    // The sample creates a three‑sheet workbook (Summary, SalesData, Notes), configures HtmlSaveOptions to export all worksheets, include hidden sheets, retain grid lines, use UTF‑8 encoding, and embed images as Base64, then saves the workbook as an IE‑compatible MHTML file. It reads the generated file, uses regular expressions to detect worksheet tab markup (either a "sheetTabs" class or anchor tags with sheet names), reports the verification result, optionally opens the .mht file with the system's default viewer (typically Internet Explorer), and finally disposes the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Path for the generated MHTML file
            string mhtmlPath = "MultiSheetWorkbook.mhtml";

            // -----------------------------------------------------------------
            // 1. Create a multi‑sheet workbook and populate it with sample data
            // -----------------------------------------------------------------
            Workbook workbook = null;
            try
            {
                workbook = new Workbook();

                // Worksheet 1
                Worksheet sheet1 = workbook.Worksheets[0];
                sheet1.Name = "Summary";
                sheet1.Cells["A1"].PutValue("Summary Sheet");
                sheet1.Cells["A2"].PutValue("Generated on:");
                sheet1.Cells["B2"].PutValue(DateTime.Now);

                // Worksheet 2
                Worksheet sheet2 = workbook.Worksheets.Add("SalesData");
                for (int row = 0; row < 10; row++)
                {
                    sheet2.Cells[row, 0].PutValue($"Product {row + 1}");
                    sheet2.Cells[row, 1].PutValue((row + 1) * 10);
                    sheet2.Cells[row, 2].PutValue((row + 1) * 5.5);
                }

                // Worksheet 3
                Worksheet sheet3 = workbook.Worksheets.Add("Notes");
                sheet3.Cells["A1"].PutValue("Notes");
                sheet3.Cells["A2"].PutValue("This workbook demonstrates MHTML export with multiple worksheets.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error while creating workbook: {ex.Message}");
                return;
            }

            // --------------------------------------------------------------
            // 2. Configure HTML save options for IE‑compatible MHTML output
            // --------------------------------------------------------------
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Export the whole workbook (all worksheets) – required for tab navigation
                ExportActiveWorksheetOnly = false,

                // Keep hidden worksheets (if any) – IE expects them to be present in the markup
                ExportHiddenWorksheet = true,

                // Export grid lines – makes the visual layout identical to Excel
                ExportGridLines = true,

                // Ensure the encoding is UTF‑8 for IE compatibility
                Encoding = Encoding.UTF8,

                // Export images as Base64 to avoid external resources
                ExportImagesAsBase64 = true
                // No need to set SaveFormat; the Save method infers it from the options
            };

            // --------------------------------------------------------------
            // 3. Save the workbook as MHTML (IE can open .mht/.mhtml files)
            // --------------------------------------------------------------
            try
            {
                workbook.Save(mhtmlPath, htmlOptions);
                Console.WriteLine($"Workbook successfully saved as MHTML to \"{mhtmlPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error while saving MHTML: {ex.Message}");
                workbook?.Dispose();
                return;
            }

            // --------------------------------------------------------------
            // 4. Verify that the generated MHTML contains worksheet tab markup
            // --------------------------------------------------------------
            try
            {
                if (!File.Exists(mhtmlPath))
                {
                    Console.WriteLine("Verification failed: MHTML file was not created.");
                    return;
                }

                // MHTML files are multipart; extract the HTML part for inspection
                string mhtmlContent = File.ReadAllText(mhtmlPath, Encoding.UTF8);

                // Simple heuristic: look for the <div> that holds the sheet tabs.
                bool tabsFound = false;

                // Pattern 1: class="sheetTabs"
                if (Regex.IsMatch(mhtmlContent, @"class\s*=\s*[""']?sheetTabs[""']?", RegexOptions.IgnoreCase))
                {
                    tabsFound = true;
                }
                else
                {
                    // Pattern 2: anchor tags with sheet names
                    string[] sheetNames = new string[workbook.Worksheets.Count];
                    for (int i = 0; i < workbook.Worksheets.Count; i++)
                        sheetNames[i] = Regex.Escape(workbook.Worksheets[i].Name);

                    string combinedPattern = string.Join("|", sheetNames);
                    if (Regex.IsMatch(mhtmlContent, $@"<a[^>]*>{combinedPattern}</a>", RegexOptions.IgnoreCase))
                    {
                        tabsFound = true;
                    }
                }

                if (tabsFound)
                {
                    Console.WriteLine("Verification succeeded: worksheet tab markup detected in the MHTML file.");
                }
                else
                {
                    Console.WriteLine("Verification failed: worksheet tab markup not found. IE may not display tabs correctly.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during verification: {ex.Message}");
            }

            // --------------------------------------------------------------
            // 5. (Optional) Launch the MHTML file using the default system handler
            // --------------------------------------------------------------
            try
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = Path.GetFullPath(mhtmlPath),
                    UseShellExecute = true
                });
                Console.WriteLine("MHTML file opened with the default viewer (typically Internet Explorer).");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unable to open MHTML file automatically: {ex.Message}");
            }

            // Clean up
            workbook?.Dispose();
        }
    }
}
