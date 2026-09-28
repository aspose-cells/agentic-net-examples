// Title: How to make a pivot table slicer printable in an Excel file using Aspose.Cells for .NET (C#)
// AI Prompts: Create a pivot table slicer and enable its IsPrintable property with Aspose.Cells for .NET. | Mark both the slicer and its underlying shape as printable so the slicer shows up on printed pages in C#.
// Common Searches: Aspose.Cells C# set slicer printable when exporting to Excel | make pivot table slicer appear on printed worksheet using Aspose.Cells | C# Aspose.Cells how to enable printing for slicer shape | set slicer IsPrintable flag Aspose.Cells .NET example
// Tags: Aspose.Cells slicer printable property | C# Aspose.Cells set slicer IsPrintable | Aspose.Cells pivot table slicer printing | Aspose.Cells shape IsPrintable for slicer | generate printable slicer Excel Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;

namespace AsposeCellsSlicerPrintableDemo
{
    // The example creates a workbook, adds sample data, builds a pivot table, inserts a slicer linked to the pivot, sets both the slicer and its shape to printable, and saves the file as SlicerPrintableDemo.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for a pivot table
            cells["A1"].PutValue("Fruit");
            cells["B1"].PutValue("Year");
            cells["C1"].PutValue("Amount");

            string[] fruits = { "Apple", "Banana", "Apple", "Banana", "Apple", "Banana" };
            int[] years = { 2020, 2020, 2021, 2021, 2022, 2022 };
            int[] amounts = { 50, 70, 60, 80, 55, 75 };

            for (int i = 0; i < fruits.Length; i++)
            {
                cells[i + 1, 0].PutValue(fruits[i]);
                cells[i + 1, 1].PutValue(years[i]);
                cells[i + 1, 2].PutValue(amounts[i]);
            }

            // Add a pivot table based on the data range
            PivotTableCollection pivots = sheet.PivotTables;
            int pivotIdx = pivots.Add("=Sheet1!A1:C7", "E2", "FruitPivot");
            PivotTable pivot = pivots[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
            pivot.AddFieldToArea(PivotFieldType.Column, "Year");
            pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a slicer linked to the pivot table
            SlicerCollection slicers = sheet.Slicers;
            int slicerIdx = slicers.Add(pivot, "G2", "Fruit");
            Slicer slicer = slicers[slicerIdx];

            // Set the slicer to be printable (using the obsolete Slicer.IsPrintable property)
            slicer.IsPrintable = true;

            // Additionally, ensure the underlying shape is printable (using Shape.IsPrintable)
            slicer.Shape.IsPrintable = true;

            // Save the workbook (lifecycle: save)
            workbook.Save("SlicerPrintableDemo.xlsx");
        }
    }
}
