// Title: Generate an Excel report with hierarchical Category/SubCategory grouping using Aspose.Cells smart markers in C#
// Common Searches: asp.net c# how to create hierarchical grouping with smart markers in Aspose.Cells | using GroupBy smart marker for nested categories in Excel with Aspose.Cells | Aspose.Cells WorkbookDesigner example for category and subcategory totals | c# code to add subtotal rows with Sum smart marker in a grouped Excel report | two‑column aggregation with GroupBy and Sum smart markers in Aspose.Cells
// Tags: Aspose.Cells GroupBy smart marker hierarchy | WorkbookDesigner hierarchical grouping Excel | C# smart marker subtotal calculation | Excel report two‑column aggregation Aspose.Cells | nested GroupBy markers for category subcategory

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, places smart markers that first group rows by Category and then by SubCategory, adds detail rows with Data markers and subtotal rows with Sum markers, supplies a DataTable of sales data, processes the template with WorkbookDesigner, and saves the hierarchical report as HierarchicalGroupingReport.xlsx.
class HierarchicalGroupingExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Access the first worksheet and name it
            var sheet = workbook.Worksheets[0];
            sheet.Name = "Report";

            // -------------------------------------------------
            // Build the template with smart markers for grouping
            // -------------------------------------------------
            // Header row
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("SubCategory");
            sheet.Cells["C1"].PutValue("Amount");

            // First level group (Category)
            sheet.Cells["A2"].PutValue("&=GroupBy(Category)");
            // Second level group (SubCategory) – nested under Category
            sheet.Cells["B3"].PutValue("&=GroupBy(SubCategory)");
            // Detail row – display the Amount value
            sheet.Cells["C4"].PutValue("&=Data(Amount)");

            // Subtotal for each SubCategory
            sheet.Cells["B5"].PutValue("SubTotal:");
            sheet.Cells["C5"].PutValue("&=Sum(Amount)");

            // Subtotal for each Category (after SubCategory group)
            sheet.Cells["A6"].PutValue("Category Total:");
            sheet.Cells["C6"].PutValue("&=Sum(Amount)");

            // Grand total (after all groups)
            sheet.Cells["A7"].PutValue("Grand Total:");
            sheet.Cells["C7"].PutValue("&=Sum(Amount)");

            // -------------------------------------------------
            // Prepare the data source
            // -------------------------------------------------
            DataTable dt = new DataTable("SalesData");
            dt.Columns.Add("Category", typeof(string));
            dt.Columns.Add("SubCategory", typeof(string));
            dt.Columns.Add("Amount", typeof(double));

            // Sample data
            dt.Rows.Add("Beverages", "Tea", 120.5);
            dt.Rows.Add("Beverages", "Coffee", 210.0);
            dt.Rows.Add("Beverages", "Juice", 85.75);
            dt.Rows.Add("Food", "Bread", 150.0);
            dt.Rows.Add("Food", "Cheese", 230.25);
            dt.Rows.Add("Food", "Fruit", 95.5);
            dt.Rows.Add("Food", "Vegetables", 180.0);
            dt.Rows.Add("Snacks", "Chips", 75.0);
            dt.Rows.Add("Snacks", "Nuts", 110.0);

            // -------------------------------------------------
            // Process the smart markers using WorkbookDesigner
            // -------------------------------------------------
            var designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(dt);
            designer.Process();

            // -------------------------------------------------
            // Save the result
            // -------------------------------------------------
            string outputPath = "HierarchicalGroupingReport.xlsx";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Report saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred while generating the report:");
            Console.WriteLine(ex.Message);
        }
    }
}
