using System;
using System.Data;
using Aspose.Cells;

namespace SmartMarkerMultiSheetDemo
{
    public class Program
    {
        public static void Main()
        {
            // Load the template workbook that contains smart markers on several worksheets
            Workbook workbook = new Workbook("TemplateWithSmartMarkers.xlsx");

            // ------------------------------------------------------------
            // Prepare a large data source (e.g., 10,000 rows) for demonstration
            // ------------------------------------------------------------
            DataTable largeTable = new DataTable("SalesData");
            largeTable.Columns.Add("Region", typeof(string));
            largeTable.Columns.Add("Product", typeof(string));
            largeTable.Columns.Add("Quantity", typeof(int));
            largeTable.Columns.Add("Revenue", typeof(double));

            // Populate the table with dummy data
            for (int i = 1; i <= 10000; i++)
            {
                largeTable.Rows.Add(
                    $"Region{(i % 5) + 1}",
                    $"Product{(i % 20) + 1}",
                    i % 100,
                    Math.Round((i % 100) * 12.34, 2));
            }

            // ------------------------------------------------------------
            // Set up the WorkbookDesigner and bind the data source
            // ------------------------------------------------------------
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook,
                // Ensure that formulas/references that span worksheets are updated after processing
                UpdateReference = true
            };

            // Bind the DataTable to the smart marker name used in the template
            designer.SetDataSource("SalesData", largeTable);

            // ------------------------------------------------------------
            // Process all smart markers in the workbook (including all worksheets)
            // ------------------------------------------------------------
            designer.Process();

            // ------------------------------------------------------------
            // Save the populated workbook
            // ------------------------------------------------------------
            workbook.Save("PopulatedMultiSheet.xlsx");
        }
    }
}