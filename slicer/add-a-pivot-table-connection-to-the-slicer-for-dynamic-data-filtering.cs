// Title: Create a pivot table and link a slicer for dynamic filtering with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add slicer to pivot table and connect for interactive filtering | how to use AddPivotConnection to bind a slicer to a pivot table in Aspose.Cells .NET | sample code for positioning slicer shape after linking to a pivot table with Aspose.Cells | C# example of creating a pivot table and slicer on the same worksheet using Aspose.Cells | dynamic data filtering with slicer and pivot table using Aspose.Cells for .NET
// Tags: Aspose.Cells add slicer to pivot table | C# pivot table slicer connection | Aspose.Cells dynamic slicer filtering | pivot table slicer shape positioning Aspose.Cells | Aspose.Cells .NET workbook pivot slicer example

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace PivotSlicerDemo
{
    // The example creates a new workbook, fills it with sample data, adds a pivot table, inserts a slicer for the 'Category' field, connects the slicer to the pivot table using AddPivotConnection for dynamic filtering, adjusts the slicer's shape properties, and saves the workbook as PivotSlicerDemo.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                Cells cells = sheet.Cells;
                cells["A1"].PutValue("Category");
                cells["B1"].PutValue("Product");
                cells["C1"].PutValue("Quantity");

                cells["A2"].PutValue("Fruit");
                cells["B2"].PutValue("Apple");
                cells["C2"].PutValue(120);

                cells["A3"].PutValue("Fruit");
                cells["B3"].PutValue("Orange");
                cells["C3"].PutValue(80);

                cells["A4"].PutValue("Vegetable");
                cells["B4"].PutValue("Carrot");
                cells["C4"].PutValue(150);

                cells["A5"].PutValue("Vegetable");
                cells["B5"].PutValue("Broccoli");
                cells["C5"].PutValue(90);

                // Define the source range for the pivot table
                string sourceRange = "A1:C5";

                // Add a pivot table (non‑obsolete overload)
                PivotTableCollection pivotTables = sheet.PivotTables;
                int pivotIndex = pivotTables.Add(sourceRange, 2, 4, "PivotTable1"); // Destination cell E3 (row 2, column 4)
                PivotTable pivotTable = pivotTables[pivotIndex];

                // Configure the pivot table fields
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Quantity");

                // Refresh and calculate the pivot data
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Add a slicer linked to the pivot table
                int slicerIndex = sheet.Slicers.Add(pivotTable, 0, 0, "CategorySlicer");
                Slicer slicer = sheet.Slicers[slicerIndex];

                // Connect the slicer to the pivot table (dynamic filtering)
                slicer.AddPivotConnection(pivotTable);

                // Set slicer position and size via its Shape object
                var shape = slicer.Shape;
                shape.Top = 200;          // distance from top of the sheet (points)
                shape.Left = 400;         // distance from left of the sheet (points)
                shape.WidthPt = 150;      // width in points
                shape.HeightPt = 200;     // height in points

                // Save the workbook to a file
                string outputPath = "PivotSlicerDemo.xlsx";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
