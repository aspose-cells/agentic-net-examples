using System;
using System.Data;
using Aspose.Cells;

namespace MultiSheetSmartMarkerDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // Worksheet 1 – Employee data smart markers
            // -------------------------------------------------
            Worksheet empSheet = workbook.Worksheets[0];
            empSheet.Name = "Employees";

            // Define smart markers in the template area (range name required when LineByLine = false)
            // Header row
            empSheet.Cells["A1"].PutValue("Name");
            empSheet.Cells["B1"].PutValue("Age");
            empSheet.Cells["C1"].PutValue("Department");

            // Data row with smart markers
            empSheet.Cells["A2"].PutValue("&=Employees.Name");
            empSheet.Cells["B2"].PutValue("&=Employees.Age");
            empSheet.Cells["C2"].PutValue("&=Employees.Department");

            // Mark the range that contains smart markers
            empSheet.Cells.CreateRange("A2:C2").Name = "_CellsSmartMarkers";

            // -------------------------------------------------
            // Worksheet 2 – Product data smart markers
            // -------------------------------------------------
            Worksheet prodSheet = workbook.Worksheets.Add("Products");

            // Header row
            prodSheet.Cells["A1"].PutValue("Product");
            prodSheet.Cells["B1"].PutValue("Price");
            prodSheet.Cells["C1"].PutValue("InStock");

            // Data row with smart markers
            prodSheet.Cells["A2"].PutValue("&=Products.ProductName");
            prodSheet.Cells["B2"].PutValue("&=Products.Price");
            prodSheet.Cells["C2"].PutValue("&=Products.InStock");

            // Mark the range that contains smart markers
            prodSheet.Cells.CreateRange("A2:C2").Name = "_CellsSmartMarkers";

            // -------------------------------------------------
            // Prepare data sources
            // -------------------------------------------------
            // Employee data table
            DataTable empTable = new DataTable("Employees");
            empTable.Columns.Add("Name", typeof(string));
            empTable.Columns.Add("Age", typeof(int));
            empTable.Columns.Add("Department", typeof(string));
            empTable.Rows.Add("John Doe", 30, "Sales");
            empTable.Rows.Add("Jane Smith", 28, "Marketing");
            empTable.Rows.Add("Bob Johnson", 45, "HR");

            // Product data table
            DataTable prodTable = new DataTable("Products");
            prodTable.Columns.Add("ProductName", typeof(string));
            prodTable.Columns.Add("Price", typeof(double));
            prodTable.Columns.Add("InStock", typeof(bool));
            prodTable.Rows.Add("Laptop", 1200.50, true);
            prodTable.Rows.Add("Smartphone", 799.99, false);
            prodTable.Rows.Add("Tablet", 450.00, true);

            // -------------------------------------------------
            // Initialize WorkbookDesigner and bind data sources
            // -------------------------------------------------
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook,
                // When using named range "_CellsSmartMarkers", set LineByLine to false
                LineByLine = false
            };

            // Bind each data source to its name used in smart markers
            designer.SetDataSource("Employees", empTable);
            designer.SetDataSource("Products", prodTable);

            // -------------------------------------------------
            // Process each worksheet individually (using Process(int, bool) overload)
            // -------------------------------------------------
            // Process Employees sheet (index 0)
            designer.Process(0, true);

            // Process Products sheet (index 1)
            designer.Process(1, true);

            // -------------------------------------------------
            // Save the resulting workbook (lifecycle rule: save)
            // -------------------------------------------------
            workbook.Save("MultiSheetSmartMarkerReport.xlsx");
        }
    }
}