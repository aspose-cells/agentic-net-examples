using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsMergeWithFormula
{
    class Program
    {
        static void Main()
        {
            // Create sample data table
            DataTable dt = new DataTable("Products");
            dt.Columns.Add("Product", typeof(string));
            dt.Columns.Add("Quantity", typeof(int));
            dt.Columns.Add("UnitPrice", typeof(double));

            dt.Rows.Add("Apple", 5, 1.2);
            dt.Rows.Add("Banana", 8, 0.8);
            dt.Rows.Add("Cherry", 12, 2.5);

            // Create a new workbook and set up the template with smart markers
            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];

            // Header row with smart markers for data binding
            sheet.Cells["A1"].PutValue("&=$Product");
            sheet.Cells["B1"].PutValue("&=$Quantity");
            sheet.Cells["C1"].PutValue("&=$UnitPrice");
            sheet.Cells["D1"].PutValue("Total");

            // First data row formula (will be repeated for each merged row)
            // B column = Quantity, C column = UnitPrice
            sheet.Cells["D2"].Formula = "=B2*C2";

            // Bind the data source and process the template
            WorkbookDesigner designer = new WorkbookDesigner(wb);
            designer.SetDataSource(dt);
            designer.Process();

            // Calculate all formulas after the merge
            wb.CalculateFormula();

            // Save the result
            wb.Save("MergedWithTotal.xlsx");
        }
    }
}