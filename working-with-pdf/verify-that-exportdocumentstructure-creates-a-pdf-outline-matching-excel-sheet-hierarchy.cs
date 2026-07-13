using System;
using Aspose.Cells;

namespace AsposeCellsPdfOutlineDemo
{
    // Author: Aspose.Cells .NET example – demonstrates ExportDocumentStructure for PDF outline
    class Program
    {
        static void Main()
        {
            // Create a new workbook and access the default worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Summary";

            // Add additional worksheets to form a hierarchy
            Worksheet sheet2 = workbook.Worksheets.Add("Details");
            Worksheet sheet3 = workbook.Worksheets.Add("Statistics");

            // Populate each sheet with sample data
            sheet1.Cells["A1"].PutValue("Report Summary");
            sheet2.Cells["A1"].PutValue("Detailed Data");
            sheet3.Cells["A1"].PutValue("Statistical Overview");

            // Create an outline (grouping) in the first sheet to illustrate hierarchy
            // Group rows 2-5 under a collapsible outline
            sheet1.Cells.GroupRows(1, 4, true);
            // Optionally set the outline to be shown when the workbook is opened
            sheet1.IsOutlineShown = true;

            // Configure PDF save options to export the document structure (outline/bookmarks)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                ExportDocumentStructure = true   // Enables PDF outline matching sheet hierarchy
            };

            // Save the workbook as PDF with the specified options
            string outputPath = "WorkbookWithOutline.pdf";
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"PDF saved to '{outputPath}' with document structure exported.");
        }
    }
}