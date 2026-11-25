// Title: Add a Fruit slicer to a pivot table and freeze rows above it using Aspose.Cells for .NET (C#)
// Common Searches: C# code to attach a slicer to a pivot table using Aspose.Cells | How to keep an Excel slicer visible while scrolling with Aspose.Cells | Programmatically freeze rows above a slicer in a .xlsx file using Aspose.Cells | Aspose.Cells example for creating a slicer for the Fruit field in a pivot table | Freeze panes at row 16 after adding a slicer with Aspose.Cells C#
// Tags: Aspose.Cells slicer‑pivot linkage C# | freeze panes after slicer insertion Aspose.Cells | C# create pivot table with slicer Aspose.Cells | Aspose.Cells set frozen rows after slicer placement | Excel slicer creation via Aspose.Cells .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace SlicerPivotFreezeDemo
{
    // The sample creates a workbook, populates it with fruit, year, and amount data, builds a pivot table, adds a slicer linked to the Fruit field, freezes rows above the slicer to keep it visible during scrolling, and saves the file as an XLSX workbook using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Path where the resulting workbook will be saved
            const string outputPath = "SlicerPivotFreezeDemo.xlsx";

            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // 2. Populate sample data that will be used for the pivot table
                //    A column – Fruit, B column – Year, C column – Amount
                sheet.Cells["A1"].PutValue("Fruit");
                sheet.Cells["B1"].PutValue("Year");
                sheet.Cells["C1"].PutValue("Amount");

                string[] fruits = { "Apple", "Banana", "Cherry", "Apple", "Banana", "Cherry", "Apple", "Banana", "Cherry" };
                int[] years =   { 2020,   2020,   2020,   2021,   2021,   2021,   2022,   2022,   2022 };
                int[] amounts = { 50,    70,    90,    60,    80,    100,   55,    75,    95 };

                for (int i = 0; i < fruits.Length; i++)
                {
                    sheet.Cells[i + 1, 0].PutValue(fruits[i]);   // Column A
                    sheet.Cells[i + 1, 1].PutValue(years[i]);    // Column B
                    sheet.Cells[i + 1, 2].PutValue(amounts[i]);  // Column C
                }

                // 3. Create a pivot table based on the data range A1:C9
                PivotTableCollection pivots = sheet.PivotTables;
                int pivotIndex = pivots.Add("=Sheet1!A1:C9", "E3", "FruitPivot");
                PivotTable pivot = pivots[pivotIndex];

                // Configure the pivot fields
                pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
                pivot.AddFieldToArea(PivotFieldType.Column, "Year");
                pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
                pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;

                // Refresh and calculate the pivot data
                pivot.RefreshData();   // Preferred over the obsolete RefreshData()
                pivot.CalculateData();

                // 4. Add a slicer linked to the pivot table for the "Fruit" field
                //    Place the slicer starting at column A (index 0) and row 13 (index 12)
                int slicerIndex = sheet.Slicers.Add(pivot, 0, 12, "FruitSlicer");
                Slicer slicer = sheet.Slicers[slicerIndex];

                // Optional: set a style for better visibility (if supported)
                // slicer.Style = SlicerStyleType.SlicerStyleLight1; // Commented out because Style property may be unavailable in some versions

                // 5. Freeze rows above the slicer so that the slicer stays visible while scrolling
                // Freeze panes just above row 16 (i.e., rows 1‑15 are frozen)
                sheet.FreezePanes("A16", 0, 0);

                // 6. Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                Console.WriteLine("The slicer is linked to the pivot table and the rows containing the slicer are frozen.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
