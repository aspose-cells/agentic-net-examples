// Title: Clear all slicers from a worksheet and save the workbook as XLSX using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that removes every slicer from a worksheet with Aspose.Cells and then saves the workbook in XLSX format. | Show how to call Worksheet.Slicers.Clear and export the workbook to an .xlsx file using Aspose.Cells for .NET.
// Common Searches: asp.net remove all slicers from an Excel sheet using Aspose.Cells | c# code to clear slicers on a worksheet and save workbook as xlsx | how to use Worksheet.Slicers.Clear method in Aspose.Cells | delete slicers programmatically with Aspose.Cells and export to XLSX
// Tags: aspocells worksheet slicers clear | c# aspocells delete slicers | aspocells save workbook xlsx | clear all slicers aspocells | worksheet slicer removal .net

using System;
using Aspose.Cells;

// Creates a workbook, clears all slicers from the first worksheet using Worksheet.Slicers.Clear, and saves the file as ClearedSlicers.xlsx in XLSX format with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Remove all slicers from the worksheet
        worksheet.Slicers.Clear();

        // Save the workbook as XLSX
        workbook.Save("ClearedSlicers.xlsx", SaveFormat.Xlsx);
    }
}
