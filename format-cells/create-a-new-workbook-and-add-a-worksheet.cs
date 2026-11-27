// Title: How to create a new Excel workbook and add a named worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells create workbook and add a sheet programmatically | set worksheet name when adding a new sheet with Aspose.Cells | save Aspose.Cells workbook as XLSX file in C# | add multiple worksheets to a newly created Excel file using Aspose.Cells .NET | remove default worksheet and insert a custom‑named sheet with Aspose.Cells
// Tags: Aspose.Cells workbook creation C# | Aspose.Cells add and rename worksheet | Aspose.Cells save workbook to XLSX | Aspose.Cells default worksheet handling | Aspose.Cells insert custom sheet C#

using Aspose.Cells;
using System;

// The example demonstrates how to instantiate a Workbook object with Aspose.Cells, add a new worksheet, assign it a custom name ("MySheet"), and save the workbook as an XLSX file using C#.
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

        // Save the workbook to verify the result (optional)
        workbook.Save("output.xlsx");
    }
}
