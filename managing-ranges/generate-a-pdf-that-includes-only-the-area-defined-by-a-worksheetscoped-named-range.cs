using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsPdfNamedRangeDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Quantity");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["B3"].PutValue(85);
                sheet.Cells["A4"].PutValue("Cherry");
                sheet.Cells["B4"].PutValue(60);

                // Define a worksheet‑scoped named range covering A1:B3
                int nameIndex = workbook.Worksheets.Names.Add("MyRange");
                // RefersTo must include the sheet name and absolute references
                workbook.Worksheets.Names[nameIndex].RefersTo = $"={sheet.Name}!$A$1:$B$3";

                // Retrieve the range object for the named range
                Name namedRange = workbook.Worksheets.Names[nameIndex];
                Aspose.Cells.Range range = namedRange.GetRange(); // fully qualified to avoid ambiguity

                // Set the print area to the address of the named range
                // This ensures that only this area is considered during PDF export
                sheet.PageSetup.PrintArea = range.Address;

                // Configure PDF save options (optional settings)
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    OnePagePerSheet = true // each sheet (or print area) on a separate PDF page
                };

                // Define output file path
                string outputPath = "NamedRangeOutput.pdf";

                // Ensure the directory for the output file exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as PDF; only the defined print area will be exported
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine("PDF generated with only the named range area.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}