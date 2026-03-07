using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace SmartMarkerLargeDataExample
{
    class Program
    {
        static void Main()
        {
            string templatePath = "TemplateWithSmartMarkers.xlsx";
            Workbook workbook;

            if (File.Exists(templatePath))
            {
                workbook = new Workbook(templatePath);
            }
            else
            {
                // Create a simple template with smart markers if it does not exist
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                ws.Cells["A1"].PutValue("&=Products.ProductID");
                ws.Cells["B1"].PutValue("&=Products.ProductName");
                ws.Cells["C1"].PutValue("&=Products.Price");
                workbook.Save(templatePath);
            }

            Worksheet sourceSheet = workbook.Worksheets[0];

            // Prepare a large data source (e.g., 5000 rows) for demonstration
            DataTable largeTable = new DataTable("Products");
            largeTable.Columns.Add("ProductID", typeof(int));
            largeTable.Columns.Add("ProductName", typeof(string));
            largeTable.Columns.Add("Price", typeof(double));

            for (int i = 1; i <= 5000; i++)
            {
                largeTable.Rows.Add(i, $"Product {i}", Math.Round(10 + i * 0.01, 2));
            }

            // Set up the WorkbookDesigner and process smart markers
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook,
                UpdateReference = true
            };

            designer.SetDataSource("Products", largeTable);
            designer.Process();

            // If the populated data exceeds a reasonable row limit,
            // move the overflow rows to a new worksheet.
            const int maxRowsPerSheet = 1000; // threshold for a single sheet

            int totalRows = sourceSheet.Cells.MaxDataRow + 1; // MaxDataRow is zero‑based

            if (totalRows > maxRowsPerSheet)
            {
                int overflowStartRow = maxRowsPerSheet; // zero‑based index where overflow begins
                int overflowRowCount = totalRows - maxRowsPerSheet;

                Worksheet overflowSheet = workbook.Worksheets.Add("OverflowData");
                overflowSheet.Cells.CopyRows(sourceSheet.Cells, overflowStartRow, 0, overflowRowCount);
                sourceSheet.Cells.DeleteRows(overflowStartRow, overflowRowCount);
            }

            workbook.Save("ProcessedLargeData.xlsx");
        }
    }
}