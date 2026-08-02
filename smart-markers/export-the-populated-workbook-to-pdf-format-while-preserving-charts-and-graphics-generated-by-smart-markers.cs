using System;
using System.Collections.Generic;
using System.Data;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfExport
{
    class Program
    {
        static void Main()
        {
            // Load a workbook that contains smart markers and charts
            Workbook workbook = new Workbook("template.xlsx");

            // Prepare a sample data source (DataTable) for smart markers
            DataTable dt = new DataTable("Products");
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Quantity", typeof(int));
            dt.Columns.Add("Price", typeof(double));

            dt.Rows.Add("Apple", 50, 0.5);
            dt.Rows.Add("Banana", 30, 0.3);
            dt.Rows.Add("Cherry", 20, 0.8);

            // Create a WorkbookDesigner, assign the workbook and set the data source
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource("Products", dt);

            // Process smart markers – this populates the worksheet and updates chart data ranges
            designer.Process();

            // Configure PDF save options to preserve document structure (charts, graphics, etc.)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            pdfOptions.ExportDocumentStructure = true;   // retain charts and graphics
            pdfOptions.CalculateFormula = true;         // ensure formulas are evaluated before saving

            // Save the populated workbook as PDF
            workbook.Save("output.pdf", pdfOptions);
        }
    }
}