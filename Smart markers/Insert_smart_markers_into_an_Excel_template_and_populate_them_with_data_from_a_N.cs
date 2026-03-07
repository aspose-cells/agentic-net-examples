using System;
using System.Data;
using Aspose.Cells;

namespace SmartMarkerDemo
{
    class Program
    {
        static void Main()
        {
            // ---------- Step 1: Prepare sample data in a DataTable ----------
            DataTable dt = new DataTable("Products");
            dt.Columns.Add("ProductID", typeof(int));
            dt.Columns.Add("ProductName", typeof(string));
            dt.Columns.Add("Price", typeof(decimal));

            dt.Rows.Add(101, "Laptop", 999.99m);
            dt.Rows.Add(102, "Smartphone", 699.49m);
            dt.Rows.Add(103, "Tablet", 399.00m);

            // ---------- Step 2: Create a template workbook with smart markers ----------
            // The workbook is created in memory; markers are placed in the first row.
            Workbook templateWb = new Workbook();
            Worksheet sheet = templateWb.Worksheets[0];

            // Smart markers use the syntax "&=TableName.ColumnName"
            sheet.Cells["A1"].PutValue("&=Products.ProductID");
            sheet.Cells["B1"].PutValue("&=Products.ProductName");
            sheet.Cells["C1"].PutValue("&=Products.Price");

            // Optional: Save the template for inspection (not required for processing)
            // templateWb.Save("Template.xlsx");

            // ---------- Step 3: Initialize WorkbookDesigner with the template ----------
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = templateWb; // Load the template workbook

            // ---------- Step 4: Set the DataTable as the data source ----------
            // The table name in the DataTable must match the smart marker prefix ("Products").
            designer.SetDataSource(dt);

            // ---------- Step 5: Process the smart markers ----------
            // This expands the markers and fills the data rows.
            designer.Process();

            // ---------- Step 6: Save the populated workbook ----------
            designer.Workbook.Save("PopulatedReport.xlsx");
        }
    }
}