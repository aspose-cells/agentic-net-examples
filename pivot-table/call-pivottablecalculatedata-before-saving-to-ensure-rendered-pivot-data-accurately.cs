// Title: Use PivotTable.CalculateData() in Aspose.Cells for .NET to compute pivot results before saving the workbook
// AI Prompts: Write C# code that creates a workbook, populates sample data, adds a pivot table, invokes PivotTable.CalculateData(), and saves the file using Aspose.Cells. | Show how to programmatically refresh pivot table calculations in an existing Excel file with Aspose.Cells before exporting it. | Demonstrate changing the source range and pivot table name while still calling CalculateData() to ensure updated pivot results.
// Common Searches: Aspose.Cells C# calculate pivot table data before workbook save | How to force pivot table refresh in Aspose.Cells .NET | PivotTable.CalculateData method example in C# | Saving Excel with calculated pivot values using Aspose.Cells | Refresh pivot cache programmatically Aspose.Cells
// Tags: Aspose.Cells PivotTable.CalculateData usage | C# calculate pivot data before saving workbook | refresh pivot cache Aspose.Cells .NET | programmatic pivot table calculation Excel | export workbook with calculated pivot Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotCalculateDataDemo
{
    // The example creates a new Workbook, fills it with sample data, adds a pivot table on the range A1:B5, calls PivotTable.CalculateData() to compute the pivot results, and saves the workbook as PivotTableWithCalculatedData.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Amount");
            sheet.Cells["A2"].PutValue("Food");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Drink");
            sheet.Cells["B3"].PutValue(80);
            sheet.Cells["A4"].PutValue("Food");
            sheet.Cells["B4"].PutValue(150);
            sheet.Cells["A5"].PutValue("Drink");
            sheet.Cells["B5"].PutValue(70);

            // Add a pivot table based on the data range A1:B5, place it at C1, and name it "SalesPivot"
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "C1", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table: Category as row field, Amount as data field
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Ensure the pivot data is calculated before saving
            pivotTable.CalculateData();

            // Save the workbook to a file
            workbook.Save("PivotTableWithCalculatedData.xlsx");
        }
    }
}
