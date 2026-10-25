// Title: Apply a custom currency number format to smart‑marker generated cells using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set custom currency format for smart marker generated cells | how to format price column as currency when using smart markers in Aspose.Cells | WorkbookDesigner apply number format to smart marker output in Excel | Aspose.Cells example of custom number format with smart markers and DataTable
// Tags: custom currency format with Aspose.Cells smart markers | WorkbookDesigner number format inheritance | smart marker cell style propagation | C# Aspose.Cells currency formatting example | Excel export smart markers custom number format

using System;
using System.Data;
using Aspose.Cells;

// The sample creates a workbook, adds a smart marker for the Price column, applies a custom currency style ($#,##0.00) to that marker cell, and then uses WorkbookDesigner to populate the sheet from a DataTable. The style is automatically inherited by all cells generated from the smart marker, producing a currency‑formatted Excel file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add column headers
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Price");

            // Insert a smart marker for the price column
            sheet.Cells["B2"].PutValue("&=Price");

            // Apply a custom currency format to the smart marker cell.
            // This style will be copied to all cells generated from the smart marker.
            Style smStyle = sheet.Cells["B2"].GetStyle();
            smStyle.Custom = "$#,##0.00"; // Currency format
            sheet.Cells["B2"].SetStyle(smStyle);

            // Prepare a data source (DataTable) with product names and prices
            DataTable data = new DataTable();
            data.Columns.Add("Product", typeof(string));
            data.Columns.Add("Price", typeof(double));
            data.Rows.Add("Apple", 1.5);
            data.Rows.Add("Banana", 0.75);
            data.Rows.Add("Cherry", 2.25);

            // Process the smart markers using WorkbookDesigner (compatible with all versions)
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(data);
            designer.Process();

            // Save the resulting workbook
            string outputPath = "SmartMarkerCurrency.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
