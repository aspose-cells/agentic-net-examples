// Title: Create a new Excel workbook and add a single named worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that uses Aspose.Cells to instantiate a Workbook, add a Worksheet called "MySheet", and save the file as an Xlsx document. | Show how to insert a worksheet into a freshly created workbook, rename it, and persist the workbook with Aspose.Cells in a .NET console application.
// Common Searches: aspnet insert worksheet into new workbook Aspose.Cells C# example | how to set worksheet name when creating Excel file with Aspose.Cells | save newly created workbook as .xlsx using Aspose.Cells | example of creating workbook with single sheet Aspose.Cells .NET
// Tags: Aspose.Cells create workbook C# | Aspose.Cells create worksheet | Aspose.Cells rename worksheet | Aspose.Cells export workbook to xlsx

using System;
using Aspose.Cells;

// // Creates a new Workbook, adds one Worksheet named "MySheet", and saves it as CreatedWorkbook.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook (default format is Xlsx)
        Workbook workbook = new Workbook();

        // Add a single worksheet to the workbook
        int newSheetIndex = workbook.Worksheets.Add();

        // Retrieve the added worksheet and optionally set its name
        Worksheet newSheet = workbook.Worksheets[newSheetIndex];
        newSheet.Name = "MySheet";

        // Save the workbook (optional, demonstrates that the workbook is valid)
        workbook.Save("CreatedWorkbook.xlsx");
    }
}
