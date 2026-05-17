using System;
using Aspose.Cells;

namespace AsposeCellsPdfExportDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (uses the provided Workbook constructor rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (including an XML‑mapped area if needed)
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");

            // If the workbook contains XML maps, they will be preserved automatically.
            // Ensure formulas are evaluated before saving.
            workbook.CalculateFormula();

            // Create PDF save options and enable document structure export
            // (preserves visual representation of mapped cells and their data)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                ExportDocumentStructure = true
            };

            // Save the workbook to PDF using the Save(string, SaveOptions) rule
            workbook.Save("MappedWorkbook.pdf", pdfOptions);

            Console.WriteLine("Workbook exported to PDF with document structure preserved.");
        }
    }
}