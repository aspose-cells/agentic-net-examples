// Title: Display monetary values with a currency symbol using smart markers in Aspose.Cells C#
// Common Searches: how to format currency values with smart markers in Aspose.Cells using C# | C# example of Aspose.Cells smart markers applying built-in currency format | using &{Amount:C} in smart markers to show currency in Excel workbook | Aspose.Cells WorkbookDesigner smart markers currency formatting tutorial | C# code to prepend currency symbol to numbers with smart markers
// Tags: WorkbookDesigner smart markers currency formatting | Aspose.Cells built-in currency format | C# Excel export monetary values | smart markers monetary formatting Aspose.Cells | Excel workbook currency symbol smart markers

using System;
using System.Data;
using Aspose.Cells;

// The example creates a new workbook, builds a DataTable with product names and amounts, inserts smart markers into cells (including the built‑in currency format &{Amount:C}), processes the markers with WorkbookDesigner, and saves the result as SmartMarkerCurrency.xlsx.
class SmartMarkerCurrencyExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Prepare sample data
            DataTable dt = new DataTable("Sales");
            dt.Columns.Add("Product", typeof(string));
            dt.Columns.Add("Amount", typeof(double));
            dt.Rows.Add("Widget", 1234.56);
            dt.Rows.Add("Gadget", 7890.12);
            dt.Rows.Add("Doohickey", 345.67);

            // Insert smart markers
            // Header row
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Revenue");
            // Data rows with smart markers
            sheet.Cells["A2"].PutValue("&=Product");
            // Use built‑in currency format instead of a custom variable
            sheet.Cells["B2"].PutValue("&{Amount:C}");

            // Use WorkbookDesigner to process smart markers
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            // Set the data source and process the smart markers
            designer.SetDataSource(dt);
            designer.Process();

            // Save the workbook
            string outputPath = "SmartMarkerCurrency.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
