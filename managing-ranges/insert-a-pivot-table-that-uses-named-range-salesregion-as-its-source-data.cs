// Title: Insert a pivot table at cell A1 from the named range 'SalesRegion' using Aspose.Cells for .NET
// AI Prompts: Write C# code that adds a pivot table named 'PivotTable1' to cell A1 of the first worksheet, using the workbook's named range 'SalesRegion' as the source, and then saves the workbook. | Generate a snippet that loads an existing Excel file with Aspose.Cells, creates a pivot table from a named range, positions it at a specific cell, and persists the changes.
// Common Searches: asp.net aspose.cells create pivot table from named range SalesRegion | c# example adding pivot table to existing workbook using named range | how to place a pivot table at A1 with Aspose.Cells .NET | using Aspose.Cells to insert pivot table and save workbook programmatically | pivot table source data from named range in Aspose.Cells C#
// Tags: Aspose.Cells add pivot table using named range | C# place pivot table at specific cell Aspose.Cells | named range as pivot source Aspose.Cells | persist workbook after pivot table insertion C# | Aspose.Cells .NET pivot table creation

using Aspose.Cells;

// // Loads input.xlsx, adds a pivot table named PivotTable1 on the first worksheet using the named range SalesRegion as the source starting at cell A1, then saves the result to output.xlsx.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Get the worksheet where the pivot table will be placed
        Worksheet sheet = workbook.Worksheets[0];

        // Insert a pivot table that uses the named range "SalesRegion" as its source data.
        // The pivot table will start at cell A1 and be named "PivotTable1".
        sheet.PivotTables.Add("=SalesRegion", "A1", "PivotTable1");

        // Save the workbook with the new pivot table
        workbook.Save("output.xlsx");
    }
}
