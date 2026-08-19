// Title: Make a Pivot Table Slicer Printable with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add sample data, build a pivot table, insert a slicer linked to the "Fruit" field, and enable printing of the slicer by setting the IsPrintable property (or the underlying Shape.IsPrintable flag) before saving the file.
// Keywords: Aspose.Cells | C# slicer printable | IsPrintable property | Shape.IsPrintable | pivot table slicer | Excel print slicer | Aspose.Cells example | generate printable slicer | .NET Excel reporting
// Common Searches: Aspose.Cells set slicer printable C# | make slicer appear on printed Excel sheet | Slicer.IsPrintable Aspose.Cells example | how to print slicer with Aspose.Cells | enable slicer printing in .NET workbook
// Developer Intent: Enable a slicer so it is included when the worksheet is printed.
// Use Cases: Create Excel reports with pivot tables where slicers must be visible on paper copies. | Automate dashboard generation that requires slicer visibility in printed handouts. | Ensure compliance documents retain interactive filter cues by printing slicers.
// AI Prompts: Generate C# code using Aspose.Cells to add a slicer to a pivot table and make it printable. | Explain the difference between Slicer.IsPrintable and Shape.IsPrintable in Aspose.Cells. | Show how to verify that a slicer appears in the printed output of an Excel file created with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;

namespace AsposeCellsSlicerPrintableDemo
{
    // Demonstrates how to create a workbook, add sample data, build a pivot table, insert a slicer linked to the "Fruit" field, and enable printing of the slicer by setting the IsPrintable property (or the underlying Shape.IsPrintable flag) before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the pivot table
            cells["A1"].Value = "Fruit";
            cells["B1"].Value = "Year";
            cells["C1"].Value = "Amount";

            string[] fruits = { "Apple", "Banana", "Cherry", "Date" };
            int[] years = { 2020, 2021 };
            int amount = 100;

            int row = 1;
            foreach (var fruit in fruits)
            {
                foreach (var year in years)
                {
                    cells[row, 0].Value = fruit;
                    cells[row, 1].Value = year;
                    cells[row, 2].Value = amount;
                    row++;
                }
            }

            // Add a pivot table based on the data range
            PivotTableCollection pivots = sheet.PivotTables;
            int pivotIdx = pivots.Add("=Sheet1!A1:C9", "E12", "FruitPivot");
            PivotTable pivot = pivots[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
            pivot.AddFieldToArea(PivotFieldType.Column, "Year");
            pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a slicer linked to the pivot table for the "Fruit" field
            SlicerCollection slicers = sheet.Slicers;
            int slicerIdx = slicers.Add(pivot, "G12", "Fruit");
            Slicer slicer = slicers[slicerIdx];

            // Set the slicer to be printable so it appears on printed pages
            slicer.IsPrintable = true; // Using the Slicer.IsPrintable property (obsolete but available)

            // Alternatively, you can also set the underlying shape's printable flag:
            // slicer.Shape.IsPrintable = true;

            // Save the workbook
            workbook.Save("SlicerPrintableDemo.xlsx");
        }
    }
}
