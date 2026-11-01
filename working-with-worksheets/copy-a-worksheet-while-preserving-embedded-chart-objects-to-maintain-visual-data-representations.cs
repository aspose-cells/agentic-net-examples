// Title: Copy an Excel worksheet while preserving embedded chart objects with Aspose.Cells for .NET
// AI Prompts: Generate C# code that copies a worksheet from a workbook, keeps all embedded chart objects intact, and saves the result to a new file using Aspose.Cells. | Write a script that duplicates a specific sheet, renames the copy, and ensures chart data bindings remain unchanged by leveraging the AddCopy method.
// Common Searches: Aspose.Cells copy worksheet with charts C# example | how to keep chart objects when duplicating an Excel sheet using .NET | AddCopy method to clone a sheet and retain chart data bindings | rename a copied worksheet programmatically with Aspose.Cells | preserve visual data representations when copying worksheets in Aspose.Cells
// Tags: duplicate worksheet preserving chart objects Aspose.Cells | Aspose.Cells AddCopy method for sheet cloning | programmatic worksheet rename after copy .NET | copy Excel sheet with embedded charts using C# | preserve visual data representations in worksheet duplication

using Aspose.Cells;
using System;

// The sample loads 'source.xlsx', selects 'Sheet1', copies the worksheet—including all embedded charts—using Worksheets.AddCopy, renames the new sheet to 'Sheet1_Copy', and saves the workbook as 'output.xlsx'.
class Program
{
    static void Main()
    {
        // Load the source workbook containing the worksheet with charts
        Workbook workbook = new Workbook("source.xlsx");

        // Access the worksheet you want to copy (by name or index)
        Worksheet sourceSheet = workbook.Worksheets["Sheet1"]; // replace with your sheet name

        // Copy the worksheet to a new position.
        // AddCopy copies the entire sheet, including all embedded chart objects,
        // preserving their visual representation and data bindings.
        int copiedSheetIndex = workbook.Worksheets.AddCopy(sourceSheet.Index);

        // Optionally rename the copied worksheet
        Worksheet copiedSheet = workbook.Worksheets[copiedSheetIndex];
        copiedSheet.Name = "Sheet1_Copy";

        // Save the workbook with the duplicated sheet and its charts
        workbook.Save("output.xlsx");
    }
}
