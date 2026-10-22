// Title: Calculate weighted average with SUMPRODUCT using Formula smart marker in Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a workbook, inserts a Formula smart marker to compute a weighted average with SUMPRODUCT on the Value and Weight columns, and processes the sheet with the designer object. | Show how to bind a DataTable containing Category, Value, and Weight fields to the workbook so that the Formula smart marker evaluates correctly for every row. | Adapt the smart marker formula to calculate a weighted median or to incorporate an extra column such as Discount in the Aspose.Cells workbook.
// Common Searches: aspnet cells smart marker sumproduct weighted average c# example | binding datatable to smart markers in aspose cells c# | c# generate excel with weighted average using aspose cells smart markers | using formula parameter in aspose cells smart markers for multiple fields | calculate weighted average in excel via aspose cells smart marker formula
// Tags: SUMPRODUCT calculation in Aspose.Cells | WorkbookDesigner data source integration | C# smart marker weighted average example | Excel generation with Aspose.Cells smart markers | DataTable to smart marker binding

using System;
using System.Data;
using Aspose.Cells;

// The C# program builds an Excel workbook, defines smart markers for Category, Value, and Weight columns, adds a Formula smart marker that uses SUMPRODUCT to calculate the weighted average, binds a multi‑row DataTable to the workbook via WorkbookDesigner, processes the markers, and saves the result as WeightedAverage.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Set column headers
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["C1"].PutValue("Weight");

            // Insert smart markers for the data rows (starting at row 2)
            sheet.Cells["A2"].PutValue("&=Category&");
            sheet.Cells["B2"].PutValue("&=Value&");
            sheet.Cells["C2"].PutValue("&=Weight&");

            // Insert a smart marker that uses the Formula parameter to calculate the weighted average
            sheet.Cells["B7"].PutValue("&=Formula=SUMPRODUCT(Weight,Value)/SUM(Weight)&");

            // Prepare a data source with multiple rows
            DataTable data = new DataTable();
            data.Columns.Add("Category", typeof(string));
            data.Columns.Add("Value", typeof(double));
            data.Columns.Add("Weight", typeof(double));

            data.Rows.Add("A", 80, 0.2);
            data.Rows.Add("B", 90, 0.3);
            data.Rows.Add("C", 70, 0.5);
            data.Rows.Add("D", 85, 0.4);
            data.Rows.Add("E", 75, 0.1);

            // Process smart markers using WorkbookDesigner (correct API)
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(data);
            designer.Process();

            // Save the resulting workbook
            string outputPath = "WeightedAverage.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
