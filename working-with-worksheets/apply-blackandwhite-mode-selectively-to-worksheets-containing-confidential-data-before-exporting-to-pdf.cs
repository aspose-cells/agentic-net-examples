using System;
using Aspose.Cells;

namespace AsposeCellsBlackAndWhiteExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add sample worksheets
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "PublicData";
            sheet1.Cells["A1"].PutValue("Public information");

            Worksheet sheet2 = workbook.Worksheets.Add("ConfidentialReport");
            sheet2.Cells["A1"].PutValue("Sensitive data");

            Worksheet sheet3 = workbook.Worksheets.Add("ConfidentialSummary");
            sheet3.Cells["A1"].PutValue("More sensitive data");

            // Apply black‑and‑white printing mode only to worksheets whose name contains "Confidential"
            foreach (Worksheet ws in workbook.Worksheets)
            {
                if (ws.Name.IndexOf("Confidential", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // Set the BlackAndWhite property to true for the worksheet's PageSetup
                    ws.PageSetup.BlackAndWhite = true;
                }
            }

            // Configure PDF save options (default options are sufficient for this scenario)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook to PDF; confidential sheets will be rendered in black and white
            workbook.Save("ConfidentialOutput.pdf", pdfOptions);
        }
    }
}