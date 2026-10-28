// Title: Add a new worksheet and create a linked PivotTable from a data range using Aspose.Cells for .NET (C#)
// Common Searches: how to add a worksheet and generate a linked pivot table with Aspose.Cells in C# | Aspose.Cells C# create pivot table on separate sheet referencing source range | sample code for linked pivot table using Aspose.Cells .NET | define source data range for pivot table in Aspose.Cells C# example | save workbook with pivot table using Aspose.Cells for .NET
// Tags: insert new worksheet Aspose.Cells .NET | generate pivot table Aspose.Cells | pivot table source range C# | pivotfieldtype row and data Aspose.Cells | save workbook with pivot table Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;

// The example creates a workbook, writes sample source data to the first sheet, adds a second sheet, defines the source range "SourceData!A1:B5", creates a linked PivotTable named "LinkedPivot" at cell C3, assigns the Category column as a row field and the Amount column as a data field, and saves the file as "LinkedPivotTable.xlsx".
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // Prepare source data on the first worksheet.
            // -------------------------------------------------
            Worksheet sourceSheet = workbook.Worksheets[0];
            sourceSheet.Name = "SourceData";

            // Header.
            sourceSheet.Cells["A1"].PutValue("Category");
            sourceSheet.Cells["B1"].PutValue("Amount");

            // Sample rows.
            sourceSheet.Cells["A2"].PutValue("Food");
            sourceSheet.Cells["B2"].PutValue(120);
            sourceSheet.Cells["A3"].PutValue("Transport");
            sourceSheet.Cells["B3"].PutValue(80);
            sourceSheet.Cells["A4"].PutValue("Food");
            sourceSheet.Cells["B4"].PutValue(150);
            sourceSheet.Cells["A5"].PutValue("Transport");
            sourceSheet.Cells["B5"].PutValue(70);

            // -------------------------------------------------
            // Add a new worksheet that will host the PivotTable.
            // -------------------------------------------------
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotSheet");

            // Define the source data range.
            string sourceData = "SourceData!A1:B5";

            // -------------------------------------------------
            // Create a PivotTable referencing the source data.
            // -------------------------------------------------
            // Add the PivotTable to the pivot sheet at cell C3.
            int pivotIndex = pivotSheet.PivotTables.Add(
                sourceData,   // source data range
                "C3",         // destination cell
                "LinkedPivot" // PivotTable name
            );

            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Add fields: Category as Row field, Amount as Data field.
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Column A (Category)
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Column B (Amount)

            // -------------------------------------------------
            // Save the workbook.
            // -------------------------------------------------
            workbook.Save("LinkedPivotTable.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
