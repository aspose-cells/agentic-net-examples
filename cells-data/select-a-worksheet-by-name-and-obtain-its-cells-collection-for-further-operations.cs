// Title: How to select a worksheet by its name and retrieve its Cells collection using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates or loads a Workbook, renames the first worksheet to a custom name, selects that worksheet by its name, and returns the Cells object for further manipulation. | Show an example of accessing a worksheet by name, obtaining its Cells collection, and writing a value to a specific cell with Aspose.Cells in a .NET application.
// Common Searches: Aspose.Cells .NET retrieve cells from a named worksheet | C# code to rename first worksheet and access it via its new name in Aspose.Cells | Write to cell A1 after selecting worksheet by custom name using Aspose.Cells | Save workbook as XLSX after modifying cells in a specific sheet with Aspose.Cells C#
// Tags: worksheet name lookup Aspose.Cells | cells collection retrieval Aspose.Cells | default sheet rename Aspose.Cells | write value to cell A1 Aspose.Cells | save workbook as xlsx Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a new Workbook, renames the default sheet to "DataSheet", selects that worksheet by its name, obtains its Cells collection, writes a string to cell A1, and saves the workbook as SelectedWorksheetDemo.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Ensure a worksheet with the desired name exists
        const string targetSheetName = "DataSheet";
        Worksheet firstSheet = workbook.Worksheets[0];
        firstSheet.Name = targetSheetName; // rename the default sheet

        // Select the worksheet by its name
        Worksheet selectedWorksheet = workbook.Worksheets[targetSheetName];

        // Obtain the Cells collection for further operations
        Cells cells = selectedWorksheet.Cells;

        // Example operation: write a value to cell A1
        cells["A1"].PutValue("Hello from the selected worksheet!");

        // Save the workbook (using the standard save method)
        workbook.Save("SelectedWorksheetDemo.xlsx", SaveFormat.Xlsx);
    }
}
