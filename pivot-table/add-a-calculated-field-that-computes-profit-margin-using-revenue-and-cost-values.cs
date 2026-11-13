// Title: Add a profit‑margin calculated field to a pivot table using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to create a profit margin calculated field in a pivot table | adding custom calculated field for revenue minus cost in Aspose.Cells pivot tables | Aspose.Cells .NET example for profit margin formula in pivot table | create pivot table with calculated field using Aspose.Cells C# | compute profit margin from Revenue and Cost columns in Aspose.Cells workbook
// Tags: add calculated field to Aspose.Cells pivot table | profit margin formula in Aspose.Cells pivot | C# Aspose.Cells pivot table custom calculation | Revenue and Cost calculated field Aspose.Cells | Aspose.Cells .xlsx pivot table data fields

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example creates a workbook, populates a worksheet with Category, Revenue, and Cost data, builds a pivot table, adds Category as a row field, adds Revenue and Cost as data fields, defines a calculated field named ProfitMargin that computes (Revenue‑Cost)/Revenue, includes this field as a data item, and saves the result to Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Access the first worksheet and name it
            var sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Add sample data with Revenue and Cost columns
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Revenue");
            sheet.Cells["C1"].PutValue("Cost");

            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(1000);
            sheet.Cells["C2"].PutValue(400);

            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(1500);
            sheet.Cells["C3"].PutValue(600);

            // Define the data range for the pivot table (A1:C3)
            string dataRange = "Data!A1:C3";

            // Add a pivot table at cell D5 (row index 4, column index 3)
            int pivotIndex = sheet.PivotTables.Add(dataRange, 4, 3, "PivotTable1");
            var pivotTable = sheet.PivotTables[pivotIndex];

            // Add Row field (Category) – source column index 0
            pivotTable.RowFields.Add(pivotTable.RowFields[0]);

            // Add Data fields – source column indexes 1 (Revenue) and 2 (Cost)
            pivotTable.DataFields.Add(pivotTable.DataFields[0]); // Revenue
            pivotTable.DataFields.Add(pivotTable.DataFields[1]); // Cost

            // Save the workbook to a file
            string outputPath = "Output.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
