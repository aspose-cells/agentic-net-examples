// Title: Retrieve the first PivotTable from the first worksheet and refresh it using Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens a workbook, accesses the first worksheet, obtains the first PivotTable from its PivotTableCollection, and calls RefreshData and CalculateData. | Show an example of iterating over a worksheet's PivotTableCollection in Aspose.Cells and performing a refresh on the pivot table at index 0. | Write a C# snippet that loads an existing Excel file, gets the first PivotTable on the first sheet, refreshes its data source, recalculates, and saves the workbook.
// Common Searches: asp.net aspose.cells get first pivot table from worksheet | c# refresh pivot table data using Aspose.Cells API | how to access PivotTableCollection and retrieve pivot at index 0 in Aspose.Cells | example code for refreshing a pivot table in Aspose.Cells for .NET | load workbook and refresh first pivot table Aspose.Cells C#
// Tags: first pivot table retrieval Aspose.Cells | pivot table refresh operation Aspose.Cells | access PivotTableCollection index Aspose.Cells | calculate pivot table data Aspose.Cells | move pivot table position Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Demonstrates creating or loading a workbook, adding sample data and a pivot table, then accessing the first worksheet's PivotTableCollection to retrieve the first PivotTable, refresh its data source, recalculate, optionally move it, and finally save the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one with new Workbook("file.xlsx"))
        Workbook workbook = new Workbook();

        // Access the first worksheet in the workbook
        Worksheet worksheet = workbook.Worksheets[0];

        // OPTIONAL: add sample data and a pivot table so that a pivot table exists.
        // This block can be removed if the workbook already contains a pivot table.
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("Fruit");
        worksheet.Cells["A3"].PutValue("Vegetable");
        worksheet.Cells["B1"].PutValue("Sales");
        worksheet.Cells["B2"].PutValue(1500);
        worksheet.Cells["B3"].PutValue(2300);

        // Add a pivot table to the worksheet (index of the new pivot table is returned)
        int pivotIndex = worksheet.PivotTables.Add("A1:B3", "D5", "SalesPivot");
        PivotTable createdPivot = worksheet.PivotTables[pivotIndex];
        createdPivot.AddFieldToArea(PivotFieldType.Row, "Category");
        createdPivot.AddFieldToArea(PivotFieldType.Data, "Sales");
        createdPivot.CalculateData();

        // Access the collection of pivot tables on the first worksheet
        PivotTableCollection pivotTables = worksheet.PivotTables;

        // Ensure there is at least one pivot table before accessing
        if (pivotTables.Count > 0)
        {
            // Obtain the first pivot table for further operations
            PivotTable firstPivotTable = pivotTables[0];

            // Example operation: refresh and recalculate the pivot table data
            firstPivotTable.RefreshData();
            firstPivotTable.CalculateData();

            // Additional operations can be performed here, e.g., moving the pivot table:
            // firstPivotTable.MoveTo(10, 2); // moves to row 10, column 2
        }

        // Save the workbook to a file
        workbook.Save("output.xlsx");
    }
}
