using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a new workbook and add sample data
                var workbook = new Workbook();
                var sheet = workbook.Worksheets[0];
                sheet.Name = "Data";
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Quantity");
                sheet.Cells["A2"].PutValue("Apples");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["A3"].PutValue("Bananas");
                sheet.Cells["B3"].PutValue(85);
                sheet.Cells["A4"].PutValue("Cherries");
                sheet.Cells["B4"].PutValue(45);

                // 2. Generate a textual summary (placeholder – replace with real AI call if available)
                string summary = "The workbook contains product quantities: Apples 120, Bananas 85, Cherries 45.";

                // 3. Write the summary to a temporary text file
                string summaryPath = Path.Combine(Path.GetTempPath(), "WorkbookSummary.txt");
                File.WriteAllText(summaryPath, summary);

                // 4. Embed the summary file as an OLE object in the worksheet
                if (File.Exists(summaryPath))
                {
                    byte[] oleData = File.ReadAllBytes(summaryPath);
                    // Rows and columns are zero‑based; using 9 (row 10) and 9 (column 10)
                    int oleIndex = sheet.OleObjects.Add(9, 9, 200, 200, oleData);
                    // Optional: display as an icon
                    sheet.OleObjects[oleIndex].DisplayAsIcon = true;
                }

                // 5. Configure PDF save options to embed OLE attachments
                var pdfOptions = new PdfSaveOptions
                {
                    EmbedAttachments = true,
                    Producer = "Aspose.Cells PDF Generator"
                };

                // 6. Save the workbook as a PDF file; the summary will be embedded
                string pdfPath = "WorkbookWithSummary.pdf";
                workbook.Save(pdfPath, pdfOptions);

                // 7. Clean up the temporary summary file
                if (File.Exists(summaryPath))
                {
                    File.Delete(summaryPath);
                }

                Console.WriteLine($"PDF saved to '{pdfPath}' with embedded summary attachment.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}