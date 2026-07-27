using System;
using System.Collections.Generic;
using System.Data;
using Aspose.Cells;

class BatchSmartMarkerProcessor
{
    static void Main()
    {
        // Define each template file, its output file, the smart‑marker name, and a distinct data source.
        var templates = new List<(string templatePath, string outputPath, string dataSourceName, DataTable data)>
        {
            ("Template1.xlsx", "Result1.xlsx", "Employees", CreateEmployeesTable()),
            ("Template2.xlsx", "Result2.xlsx", "Products",  CreateProductsTable())
        };

        foreach (var item in templates)
        {
            // Load the template workbook (Workbook(string) constructor).
            Workbook workbook = new Workbook(item.templatePath);

            // Create a WorkbookDesigner for this workbook (WorkbookDesigner(Workbook) constructor).
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Assign the distinct data source to the designer (SetDataSource(string, object) method).
            designer.SetDataSource(item.dataSourceName, item.data);

            // Process the smart markers (Process() method).
            designer.Process();

            // Save the processed workbook (Workbook.Save(string) method).
            designer.Workbook.Save(item.outputPath);
        }
    }

    // Sample data source for the first template.
    private static DataTable CreateEmployeesTable()
    {
        DataTable dt = new DataTable("Employees");
        dt.Columns.Add("Name", typeof(string));
        dt.Columns.Add("Age",  typeof(int));
        dt.Rows.Add("John Doe",   30);
        dt.Rows.Add("Jane Smith", 28);
        return dt;
    }

    // Sample data source for the second template.
    private static DataTable CreateProductsTable()
    {
        DataTable dt = new DataTable("Products");
        dt.Columns.Add("ProductID",   typeof(int));
        dt.Columns.Add("ProductName", typeof(string));
        dt.Columns.Add("Price",       typeof(decimal));
        dt.Rows.Add(1, "Laptop",     1200.00m);
        dt.Rows.Add(2, "Smartphone",  799.99m);
        return dt;
    }
}