// Title: Create a slicer linked to a pivot table in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: c# asp.net add slicer to a pivot table with Aspose.Cells | how to programmatically create an Excel slicer linked to a pivot table using Aspose.Cells | sample code for linking a slicer to a pivot table in a .xlsx file with C# | Aspose.Cells example to insert a slicer for a pivot table and save workbook | create and save Excel workbook with slicer for pivot table using .NET API
// Tags: Aspose.Cells add slicer to pivot table | C# create Excel slicer programmatically | Aspose.Cells slicer API | Excel pivot table slicer .NET | Save workbook with slicer Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerDemo
{
    // The example creates a new workbook, populates it with sample fruit, year, and amount data, builds a pivot table, applies a style, refreshes and calculates the pivot data, adds a slicer linked to the pivot table, and saves the workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data (A1:C9)
                cells["A1"].PutValue("Fruit");
                cells["B1"].PutValue("Year");
                cells["C1"].PutValue("Amount");

                string[] fruits = { "Grape", "Blueberry", "Kiwi", "Cherry", "Grape", "Blueberry", "Kiwi", "Cherry" };
                int[] years =   { 2020,   2020,      2020,   2020,    2021,   2021,       2021,   2021 };
                int[] amounts = { 50,    60,        70,     80,      90,     100,        110,    120 };

                for (int i = 0; i < fruits.Length; i++)
                {
                    cells[i + 2, 0].PutValue(fruits[i]);   // Column A
                    cells[i + 2, 1].PutValue(years[i]);   // Column B
                    cells[i + 2, 2].PutValue(amounts[i]); // Column C
                }

                // Create a pivot table based on the data range A1:C9
                // Use the overload that does not require NewPivotTableOptions
                int pivotIndex = sheet.PivotTables.Add("=Sheet1!A1:C9", 2, 4, "FruitPivot");
                PivotTable pivot = sheet.PivotTables[pivotIndex];

                // Add fields to the pivot table
                pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
                pivot.AddFieldToArea(PivotFieldType.Column, "Year");
                pivot.AddFieldToArea(PivotFieldType.Data, "Amount");

                // Apply a built‑in style (optional)
                pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium10;

                // Refresh and calculate the pivot table data
                pivot.RefreshData();   // correct method
                pivot.CalculateData();

                // Add a slicer linked to the pivot table.
                int slicerIndex = sheet.Slicers.Add(pivot, 0, 0, "FruitSlicer");
                Slicer slicer = sheet.Slicers[slicerIndex];

                // Save the workbook with the slicer attached
                string outputPath = "Workbook_With_Slicer.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
