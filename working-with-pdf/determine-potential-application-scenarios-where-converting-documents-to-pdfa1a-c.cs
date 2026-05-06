using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace PdfA1aComplianceDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some sample data.
            // This data could represent any type of document that needs long‑term preservation.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "ArchiveData";

            // Example content: financial report header.
            sheet.Cells["A1"].PutValue("Financial Report 2023");
            sheet.Cells["A2"].PutValue("Prepared By");
            sheet.Cells["B2"].PutValue("Acme Corp");
            sheet.Cells["A3"].PutValue("Date");
            sheet.Cells["B3"].PutValue(DateTime.Now.ToString("yyyy-MM-dd"));
            sheet.Cells["A5"].PutValue("Revenue");
            sheet.Cells["B5"].PutValue(1250000);
            sheet.Cells["A6"].PutValue("Expenses");
            sheet.Cells["B6"].PutValue(830000);
            sheet.Cells["A7"].PutValue("Net Profit");
            sheet.Cells["B7"].PutValue(420000);

            // ------------------------------------------------------------
            // Set PDF/A‑1a compliance.
            // PDF/A‑1a (ISO 19005‑1) preserves both visual appearance and
            // document structure (tags, metadata). This is required for:
            //   • Legal contracts that must remain readable and searchable.
            //   • Medical records subject to health‑care regulations.
            //   • Financial statements for audit trails.
            //   • Government documents that must meet archival standards.
            //   • Any content that needs guaranteed long‑term accessibility.
            // ------------------------------------------------------------
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Enforce PDF/A‑1a compliance.
                Compliance = PdfCompliance.PdfA1a,

                // Optional: embed all fonts to ensure visual fidelity.
                EmbedStandardWindowsFonts = true,

                // Optional: export document structure for better accessibility.
                ExportDocumentStructure = true
            };

            // Save the workbook as a PDF/A‑1a compliant file.
            string outputPath = "FinancialReport_2023_PdfA1a.pdf";
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"PDF/A‑1a document generated: {outputPath}");
        }
    }
}