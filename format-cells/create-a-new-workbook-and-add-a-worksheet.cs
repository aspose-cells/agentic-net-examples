// Title: How to create a new Excel workbook and add a named worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# create workbook and add a sheet with custom name | C# code to add a new worksheet to an existing Aspose.Cells workbook | Saving a new Excel file with a specific sheet name using Aspose.Cells .NET | Initialize Aspose.Cells workbook and programmatically add a worksheet in C# | Example of adding a worksheet to a workbook and saving as .xlsx with Aspose.Cells
// Tags: create workbook Aspose.Cells .NET | add worksheet Aspose.Cells C# | rename worksheet Aspose.Cells | save workbook to xlsx Aspose.Cells | initialize Aspose.Cells workbook

using System;
using Aspose.Cells;

// The example shows how to instantiate a new Workbook with Aspose.Cells, add a new Worksheet named "MySheet", and save the file as "output.xlsx" using C#.
class Program
{
    static void Main()
    {
        // Create a new workbook (contains a default worksheet)
        Workbook workbook = new Workbook();

        // Add a new worksheet to the workbook
        int newSheetIndex = workbook.Worksheets.Add();
        Worksheet newSheet = workbook.Worksheets[newSheetIndex];
        newSheet.Name = "MySheet";

        // (Optional) Save the workbook to verify the result
        workbook.Save("output.xlsx");
    }
}
