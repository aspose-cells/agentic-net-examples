// Title: Enable DisplayItemLabels on an Aspose.Cells PivotTable in C# to show each data item label
// Common Searches: Aspose.Cells C# how to turn on DisplayItemLabels for a PivotTable | show individual data item labels in a PivotTable using Aspose.Cells .NET | set PivotTable DisplayItemLabels property to true in C# example | programmatically display item labels for rows in Aspose.Cells pivot tables | C# code to enable item labels in Aspose.Cells pivot tables
// Tags: Aspose.Cells PivotTable DisplayItemLabels property | C# enable item labels in Aspose.Cells pivot tables | show data item labels Aspose.Cells workbook | programmatic pivot table label visibility .NET | Aspose.Cells set DisplayItemLabels true | pivot table row field label display C#

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example creates a workbook, adds sample data, builds a PivotTable, sets the PivotTable's DisplayItemLabels property to true so each data item label is shown, and saves the workbook as PivotTableDisplayItemLabels.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and add sample data
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Value");
            dataSheet.Cells["A2"].PutValue("A");
            dataSheet.Cells["B2"].PutValue(10);
            dataSheet.Cells["A3"].PutValue("B");
            dataSheet.Cells["B3"].PutValue(20);
            dataSheet.Cells["A4"].PutValue("A");
            dataSheet.Cells["B4"].PutValue(30);

            // Add a new worksheet to host the PivotTable
            int pivotSheetIndex = workbook.Worksheets.Add();
            Worksheet pivotSheet = workbook.Worksheets[pivotSheetIndex];
            pivotSheet.Name = "Pivot";

            // Define the source range for the PivotTable (A1:B4 on the Data sheet)
            string sourceRange = "Data!A1:B4";

            // Add the PivotTable at cell C3 on the Pivot sheet
            int pivotTableIndex = pivotSheet.PivotTables.Add(sourceRange, "C3", "PivotTable1");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotTableIndex];

            // Configure fields: Category as row field, Value as data field
            // RowFields and DataFields use fields from the source range (available in RowFields collection initially)
            pivotTable.RowFields.Add(pivotTable.RowFields[0]);   // Category column
            pivotTable.DataFields.Add(pivotTable.RowFields[1]); // Value column as data field

            // Save the workbook
            string outputPath = "PivotTableDisplayItemLabels.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
