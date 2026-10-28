// Title: Add Region and Product Category as row fields to an existing Excel pivot table with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells add region row field to existing pivot table | How to group sales data by product category in a pivot table using Aspose.Cells | Insert multiple row fields into an Excel pivot table programmatically with Aspose.Cells .NET | Aspose.Cells pivot table row field example for sales report | Modify existing pivot table rows in C# using Aspose.Cells library
// Tags: add row fields to pivot table Aspose.Cells | group sales data by region Aspose.Cells | product category row field C# | modify existing pivot table .NET | Excel pivot table row grouping Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Required for PivotTable class

namespace AsposeCellsExample
{
    // The C# program loads 'SalesData.xlsx', verifies a pivot table exists on the first worksheet, adds the first two source columns (Region and Product Category) as row fields to the pivot table, and saves the updated workbook as 'SalesData_Grouped.xlsx', handling missing files and operation errors.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "SalesData.xlsx";
            const string outputPath = "SalesData_Grouped.xlsx";

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook containing the sales data
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure there is at least one pivot table
                if (worksheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("No pivot tables found in the worksheet.");
                    return;
                }

                // Retrieve the first pivot table on the worksheet
                PivotTable pivotTable = worksheet.PivotTables[0];

                try
                {
                    // Add row fields to group data by Region (first column in source data)
                    pivotTable.AddFieldToArea(PivotFieldType.Row, 0); // Region

                    // Add row fields to group data by Product Category (second column in source data)
                    pivotTable.AddFieldToArea(PivotFieldType.Row, 1); // Product Category
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error configuring pivot table: {ex.Message}");
                    return;
                }

                try
                {
                    // Save the modified workbook
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error saving workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
