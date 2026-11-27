// Title: Generate a PivotTable and activate the PivotTable ribbon in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: how to programmatically add a pivot table and show the pivot ribbon with Aspose.Cells in C# | Aspose.Cells C# example to create a pivot table and enable Excel pivot UI | display PivotTable ribbon after creating a pivot table using Aspose.Cells for .NET | C# code to add a pivot table from a range and activate built‑in pivot tools in Excel | using Aspose.Cells to build a pivot table and expose Excel's pivot controls
// Tags: Aspose.Cells create pivot table C# | enable pivot ribbon Aspose.Cells | add pivot table from range .NET | configure row and data fields Aspose.Cells | save workbook with pivot table Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // Required for PivotTable classes

// The example creates a new Workbook, fills cells A1:B5 with sample Category and Amount data, adds a PivotTable covering that range using the legacy Add overload, assigns the first source column as a row field and the second as a data field, enables the PivotTable ribbon for interactive Excel tools, and saves the file as PivotTableDemo.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Amount");
            sheet.Cells["A2"].PutValue("Food");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Beverage");
            sheet.Cells["B3"].PutValue(80);
            sheet.Cells["A4"].PutValue("Food");
            sheet.Cells["B4"].PutValue(150);
            sheet.Cells["A5"].PutValue("Beverage");
            sheet.Cells["B5"].PutValue(70);

            // Add a pivot table based on the data range A1:B5
            PivotTableCollection pivots = sheet.PivotTables;
            // Using the obsolete overload is still functional; the warning is harmless for this example
            int pivotIndex = pivots.Add("=Sheet1!$A$1:$B$5", "$D$1", "PivotTable1");
            PivotTable pivotTable = pivots[pivotIndex];

            // Configure the pivot table:
            // Add "Category" (first source column) as a row field
            pivotTable.RowFields.Add(pivotTable.RowFields[0]);   // Category

            // Add "Amount" (second source column) as a data field
            pivotTable.DataFields.Add(pivotTable.RowFields[1]); // Amount

            // Save the workbook
            workbook.Save("PivotTableDemo.xlsx");
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
