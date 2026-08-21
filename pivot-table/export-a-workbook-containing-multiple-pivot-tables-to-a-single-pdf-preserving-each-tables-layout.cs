// Title: Export Multiple Pivot Tables to a Single PDF with Aspose.Cells (.NET)
// Description: Create a workbook with two data sheets, add a tabular and an outline pivot table, refresh them, set PdfSaveOptions (ExportDocumentStructure, OnePagePerSheet = false), and save the entire workbook as one PDF while preserving each pivot's layout and sheet bookmarks.
// Keywords: Aspose.Cells | C# | .NET | PivotTable | PDF export | multiple pivot tables | PdfSaveOptions | ExportDocumentStructure | preserve layout | single PDF
// Common Searches: Aspose.Cells export multiple pivot tables to PDF | save workbook with several pivot tables as one PDF | preserve pivot table formatting when converting to PDF | PdfSaveOptions ExportDocumentStructure example | C# export pivot tables to PDF using Aspose.Cells
// Developer Intent: Generate one PDF that contains all pivot tables from a workbook, keeping each table’s formatting intact.
// Use Cases: Combine category‑wise and region‑wise pivot tables into a single PDF report for executive review. | Automate monthly financial statements where each department’s pivot table appears on its own sheet with bookmarks. | Produce a printable summary that includes multiple analytical pivot tables without losing tabular or outline layouts.
// AI Prompts: Show how to insert page breaks between pivot tables when exporting to PDF with Aspose.Cells. | Provide code to add PDF metadata (title, author, subject) while saving multiple pivot tables to a single file. | Explain how to customize PDF bookmarks to match the names of pivot‑table worksheets.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Saving;

namespace AsposeCellsPivotPdfExport
{
    // Create a workbook with two data sheets, add a tabular and an outline pivot table, refresh them, set PdfSaveOptions (ExportDocumentStructure, OnePagePerSheet = false), and save the entire workbook as one PDF while preserving each pivot's layout and sheet bookmarks.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // Sheet 1 – Sample data and first pivot table
            // -------------------------------------------------
            Worksheet dataSheet1 = workbook.Worksheets[0];
            dataSheet1.Name = "Data1";

            // Populate sample data
            dataSheet1.Cells["A1"].PutValue("Category");
            dataSheet1.Cells["B1"].PutValue("Amount");
            dataSheet1.Cells["A2"].PutValue("Food");
            dataSheet1.Cells["B2"].PutValue(1200);
            dataSheet1.Cells["A3"].PutValue("Beverage");
            dataSheet1.Cells["B3"].PutValue(800);
            dataSheet1.Cells["A4"].PutValue("Stationery");
            dataSheet1.Cells["B4"].PutValue(400);

            // Add a worksheet to host the first pivot table
            Worksheet pivotSheet1 = workbook.Worksheets.Add("Pivot1");
            int ptIndex1 = pivotSheet1.PivotTables.Add("=Data1!A1:B4", "A3", "PivotTable1");
            PivotTable pivotTable1 = pivotSheet1.PivotTables[ptIndex1];
            pivotTable1.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable1.AddFieldToArea(PivotFieldType.Data, "Amount");
            pivotTable1.ShowInTabularForm(); // Preserve tabular layout

            // -------------------------------------------------
            // Sheet 2 – Sample data and second pivot table
            // -------------------------------------------------
            Worksheet dataSheet2 = workbook.Worksheets.Add("Data2");
            dataSheet2.Cells["A1"].PutValue("Region");
            dataSheet2.Cells["B1"].PutValue("Sales");
            dataSheet2.Cells["A2"].PutValue("North");
            dataSheet2.Cells["B2"].PutValue(1500);
            dataSheet2.Cells["A3"].PutValue("South");
            dataSheet2.Cells["B3"].PutValue(1700);
            dataSheet2.Cells["A4"].PutValue("East");
            dataSheet2.Cells["B4"].PutValue(1300);
            dataSheet2.Cells["A5"].PutValue("West");
            dataSheet2.Cells["B5"].PutValue(1100);

            // Add a worksheet to host the second pivot table
            Worksheet pivotSheet2 = workbook.Worksheets.Add("Pivot2");
            int ptIndex2 = pivotSheet2.PivotTables.Add("=Data2!A1:B5", "A3", "PivotTable2");
            PivotTable pivotTable2 = pivotSheet2.PivotTables[ptIndex2];
            pivotTable2.AddFieldToArea(PivotFieldType.Row, "Region");
            pivotTable2.AddFieldToArea(PivotFieldType.Data, "Sales");
            pivotTable2.ShowInOutlineForm(); // Preserve outline layout

            // -------------------------------------------------
            // Refresh all pivot tables to ensure they contain up‑to‑date data
            // -------------------------------------------------
            workbook.Worksheets.RefreshPivotTables();

            // -------------------------------------------------
            // Configure PDF save options
            // -------------------------------------------------
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Preserve the document structure (bookmarks for each sheet)
                ExportDocumentStructure = true,

                // Keep each sheet on its own page(s) – default behavior
                OnePagePerSheet = false,

                // Do not force all columns onto a single page
                AllColumnsInOnePagePerSheet = false
            };

            // -------------------------------------------------
            // Save the workbook (containing both pivot tables) to a single PDF file
            // -------------------------------------------------
            workbook.Save("MultiplePivotTables.pdf", pdfOptions);

            Console.WriteLine("Workbook with multiple pivot tables exported to PDF successfully.");
        }
    }
}
