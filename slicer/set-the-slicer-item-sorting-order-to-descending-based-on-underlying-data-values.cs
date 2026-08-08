// Title: Set Slicer Items to Descending Order by Data Values in Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, build a pivot table, add a slicer linked to a row field, and configure the slicer to sort its items in descending order based on the underlying data values using Aspose.Cells for C#.
// Keywords: Aspose.Cells slicer descending sort | C# slicer sort order | Aspose.Cells pivot table slicer | SortOrder.Descending Aspose.Cells | slicer item sorting by value | .NET Excel slicer example
// Common Searches: Aspose.Cells how to sort slicer items descending | C# set slicer sort order based on data values | pivot table slicer descending order Aspose.Cells | change slicer sort order .NET Excel | descending slicer items Aspose.Cells example
// Developer Intent: Configure a slicer so its items appear in descending order according to the values in the linked pivot table.
// Use Cases: Generate a financial report where the slicer lists categories from highest to lowest amount. | Create an interactive dashboard that ranks items automatically by descending values. | Align slicer ordering with pivot table auto‑sorting for consistent data presentation.
// AI Prompts: Write C# code with Aspose.Cells to add a slicer to a pivot table and set its items to sort descending by the data field. | Explain the effect of SortOrder.Descending on slicer items and its relationship with pivot field auto‑sorting. | Provide a step‑by‑step tutorial for building a pivot table, attaching a slicer, and applying descending sort order in a .NET application.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerSortDemo
{
    // Demonstrates how to create a workbook, build a pivot table, add a slicer linked to a row field, and configure the slicer to sort its items in descending order based on the underlying data values using Aspose.Cells for C#.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Amount");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["B3"].PutValue(300);
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B4"].PutValue(150);
                sheet.Cells["A5"].PutValue("D");
                sheet.Cells["B5"].PutValue(80);

                // Create a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B5", "D1", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add the Category field as a row field and Amount as a data field
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

                // Ensure the pivot field is sorted descending
                PivotField rowField = pivotTable.RowFields[0];
                rowField.IsAutoSort = true;      // enable auto‑sorting
                rowField.IsAscendSort = false;   // descending order

                // Add a slicer linked to the pivot table for the Category field
                // Note: Add method signature is (PivotTable, firstRow, firstColumn, fieldName)
                int slicerIndex = sheet.Slicers.Add(pivotTable, 0, 5, "Category");
                Slicer slicer = sheet.Slicers[slicerIndex];

                // Set the slicer items to be sorted in descending order based on the underlying data values
                slicer.SortOrderType = SortOrder.Descending;

                // Save the workbook
                workbook.Save("SlicerSortedDesc.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
