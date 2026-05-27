using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsFitToPagesTallDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle: create)
                Workbook workbook = new Workbook();

                // ---------- Sheet 1 ----------
                Worksheet sheet1 = workbook.Worksheets[0];
                sheet1.Name = "FirstSheet";

                // Fill sample data (enough rows to span multiple pages vertically)
                for (int i = 0; i < 120; i++)
                    for (int j = 0; j < 8; j++)
                        sheet1.Cells[i, j].PutValue($"R{i + 1}C{j + 1}");

                // Configure page setup: keep FitToPagesWide unchanged, set FitToPagesTall = 1
                PageSetup ps1 = sheet1.PageSetup;
                ps1.PrintArea = "A1:H120";
                ps1.FitToPagesTall = 1; // force all rows onto one page tall

                // ---------- Sheet 2 ----------
                Worksheet sheet2 = workbook.Worksheets.Add("SecondSheet");

                // Fill sample data
                for (int i = 0; i < 80; i++)
                    for (int j = 0; j < 12; j++)
                        sheet2.Cells[i, j].PutValue($"R{i + 1}C{j + 1}");

                // Configure page setup similarly
                PageSetup ps2 = sheet2.PageSetup;
                ps2.PrintArea = "A1:L80";
                ps2.FitToPagesTall = 1; // same scaling rule as sheet1

                // Set PDF save options – scaling is driven by the page setup settings
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    OnePagePerSheet = false,               // allow multiple pages if needed
                    AllColumnsInOnePagePerSheet = false    // keep column scaling per sheet
                };

                // Save the workbook as a multi‑sheet PDF (lifecycle: save)
                string outputPath = "MultiSheetOutput.pdf";

                // Ensure we can write to the output location
                try
                {
                    if (File.Exists(outputPath))
                        File.Delete(outputPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Could not delete existing file '{outputPath}'. {ex.Message}");
                }

                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine("PDF generated with FitToPagesTall = 1 on each sheet.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}