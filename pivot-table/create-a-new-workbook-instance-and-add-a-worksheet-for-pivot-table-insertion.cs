// Title: How to create a new Aspose.Cells Workbook and add a named worksheet for pivot table data in C#
// Common Searches: C# Aspose.Cells create workbook and add a sheet named PivotData for a pivot table | How to add a worksheet to a new Aspose.Cells workbook before generating a pivot table | Initialize an Aspose.Cells workbook and set worksheet name for pivot table source data in .NET | Aspose.Cells example: create workbook, add worksheet, rename to PivotData for pivot table insertion | Programmatically create a workbook and worksheet for pivot table data using Aspose.Cells in C#
// Tags: Aspose.Cells workbook initialization C# | worksheet addition for pivot tables Aspose.Cells | rename worksheet to PivotData Aspose.Cells | prepare worksheet source data Aspose.Cells | Aspose.Cells pivot table worksheet setup

using System;
using Aspose.Cells;

// The sample demonstrates how to instantiate a new Aspose.Cells Workbook in C#, add a worksheet, retrieve its index, rename the sheet to "PivotData", and optionally save the workbook, all wrapped in error handling.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook instance
            Workbook workbook = new Workbook();

            // Add a new worksheet and get its index
            int sheetIndex = workbook.Worksheets.Add();

            // Get a reference to the newly added worksheet
            Worksheet pivotWorksheet = workbook.Worksheets[sheetIndex];

            // Set the worksheet name to "PivotData"
            pivotWorksheet.Name = "PivotData";

            // (Optional) Save the workbook if you need to persist it
            // workbook.Save("PivotWorkbook.xlsx");
        }
        catch (Exception ex)
        {
            // Log or handle exceptions as needed
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
