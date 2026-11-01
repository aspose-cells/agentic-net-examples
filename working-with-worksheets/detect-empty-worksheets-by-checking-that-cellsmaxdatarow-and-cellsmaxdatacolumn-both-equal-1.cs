// Title: Identify empty worksheets in an Excel file with Aspose.Cells for .NET by checking MaxDataRow and MaxDataColumn
// AI Prompts: Generate C# code using Aspose.Cells that iterates through all worksheets and prints the names of sheets where Cells.MaxDataRow and Cells.MaxDataColumn are both -1. | Write a reusable C# method that receives a Workbook object and returns a List<string> of worksheet names that have no data, based on the -1 values of MaxDataRow and MaxDataColumn. | Adapt the example to log the empty‑sheet detection results to a text file instead of the console and skip saving the workbook when no changes are made.
// Common Searches: Aspose.Cells how to programmatically determine if a worksheet contains any data in C# | C# check for empty sheets in an Excel workbook using MaxDataRow and MaxDataColumn | list worksheets with no cells populated using Aspose.Cells .NET | detect blank worksheets in Excel file with Aspose.Cells API
// Tags: detect empty worksheets Aspose.Cells .NET | Cells.MaxDataRow MaxDataColumn check | list blank sheets C# Aspose.Cells | worksheet data presence detection Excel .NET | empty sheet detection using Aspose.Cells API

using System;
using Aspose.Cells;

// The sample loads an Excel workbook, iterates over each worksheet, and uses Cells.MaxDataRow and Cells.MaxDataColumn (both -1 indicate no data) to determine whether the sheet is empty, outputting the result for each worksheet.
class DetectEmptyWorksheets
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // MaxDataRow and MaxDataColumn return -1 when the sheet has no data
            bool isEmpty = sheet.Cells.MaxDataRow == -1 && sheet.Cells.MaxDataColumn == -1;

            // Output the result for each worksheet
            Console.WriteLine($"Worksheet \"{sheet.Name}\" is {(isEmpty ? "empty" : "not empty")}.");
        }

        // (Optional) Save the workbook if any modifications were made
        // workbook.Save("output.xlsx");
    }
}
