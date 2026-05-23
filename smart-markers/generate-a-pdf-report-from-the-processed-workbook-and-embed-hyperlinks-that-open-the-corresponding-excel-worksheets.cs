using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class PdfReportWithHyperlinks
{
    static void Main()
    {
        try
        {
            // Create a new workbook (uses Workbook() constructor rule)
            using (Workbook workbook = new Workbook())
            {
                // Rename the default sheet to "Summary"
                workbook.Worksheets[0].Name = "Summary";

                // Add a few sample worksheets with some data
                int numberOfSheets = 3;
                for (int i = 1; i <= numberOfSheets; i++)
                {
                    // Add a new worksheet with a specific name; Add(string) returns Worksheet
                    Worksheet ws = workbook.Worksheets.Add($"Sheet{i}");
                    ws.Cells["A1"].PutValue($"Content of {ws.Name}");
                }

                // Insert hyperlinks in the Summary sheet that point to each worksheet
                Worksheet summarySheet = workbook.Worksheets["Summary"];
                for (int i = 1; i <= numberOfSheets; i++)
                {
                    Worksheet targetSheet = workbook.Worksheets[i];
                    int row = i; // start from row 1 (0‑based index)

                    // Display text in the cell
                    summarySheet.Cells[row, 0].PutValue($"Go to {targetSheet.Name}");

                    // Create an internal hyperlink to the target sheet's cell A1
                    // Address format: "#SheetName!A1"
                    string hyperlinkAddress = $"#{targetSheet.Name}!A1";
                    summarySheet.Hyperlinks.Add(row, 0, 1, 1, hyperlinkAddress);
                }

                // Configure PDF save options (uses PdfSaveOptions property rules)
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    EmbedAttachments = false,               // optional, demonstrates property usage
                    ExportDocumentStructure = true          // retain document structure for accessibility
                };

                // Determine output path and ensure the directory exists
                string outputPath = "Report.pdf";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as a PDF with the embedded hyperlinks
                // (uses Workbook.Save(string, SaveOptions) rule)
                workbook.Save(outputPath, pdfOptions);
            }
        }
        catch (FileNotFoundException fnfEx)
        {
            Console.Error.WriteLine($"File not found: {fnfEx.FileName}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}