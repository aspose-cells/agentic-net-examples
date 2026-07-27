using System;
using System.Data;
using Aspose.Cells;
using Aspose.Cells.Tables;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // ----- Smart marker template (hierarchical grouping) -----
        // Header
        sheet.Cells["A1"].PutValue("Two‑Column Aggregation Report");

        // Start grouping by Category
        sheet.Cells["A3"].PutValue("{{#group Category}}");
        // Display Category name
        sheet.Cells["A4"].PutValue("Category: {{Category}}");

        // Nested grouping by SubCategory
        sheet.Cells["A5"].PutValue("{{#group SubCategory}}");
        // Display SubCategory and aggregated columns (Sales and Quantity)
        sheet.Cells["A6"].PutValue("SubCategory: {{SubCategory}}");
        sheet.Cells["B6"].PutValue("Sales: {{Sales}}");
        sheet.Cells["C6"].PutValue("Quantity: {{Quantity}}");
        // End SubCategory group
        sheet.Cells["A7"].PutValue("{{/group}}");

        // End Category group
        sheet.Cells["A8"].PutValue("{{/group}}");

        // ----- Prepare data source -----
        DataTable dt = new DataTable("Data");
        dt.Columns.Add("Category", typeof(string));
        dt.Columns.Add("SubCategory", typeof(string));
        dt.Columns.Add("Sales", typeof(double));
        dt.Columns.Add("Quantity", typeof(int));

        // Sample rows
        dt.Rows.Add("Fruit", "Apple", 1200.5, 10);
        dt.Rows.Add("Fruit", "Apple", 800.0, 5);
        dt.Rows.Add("Fruit", "Banana", 500.0, 8);
        dt.Rows.Add("Vegetable", "Carrot", 300.0, 12);
        dt.Rows.Add("Vegetable", "Carrot", 450.0, 7);
        dt.Rows.Add("Vegetable", "Potato", 600.0, 15);

        // ----- Process smart markers -----
        WorkbookDesigner designer = new WorkbookDesigner();
        designer.Workbook = workbook;
        designer.SetDataSource(dt);
        designer.Process();

        // Save the resulting workbook
        workbook.Save("HierarchicalGroupingSmartMarker.xlsx");
    }
}