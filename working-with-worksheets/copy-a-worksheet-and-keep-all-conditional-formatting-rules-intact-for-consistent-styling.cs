// Title: Copy an Excel worksheet while preserving all conditional formatting rules using Aspose.Cells for .NET
// AI Prompts: Duplicate a worksheet in the same workbook and keep every conditional formatting rule unchanged with Aspose.Cells in C#. | Create a copy of a specific sheet, retain its conditional formatting, rename it, and save the workbook using Aspose.Cells for .NET. | Clone an Excel sheet, ensure all conditional formatting is transferred, and write the result to a new file with Aspose.Cells.
// Common Searches: how to copy worksheet with conditional formatting using Aspose.Cells C# | preserve conditional formatting when duplicating an Excel sheet in .NET | Aspose.Cells copy sheet retain conditional formatting rules
// Tags: copy sheet with conditional formatting Aspose.Cells | preserve formatting rules on worksheet clone .NET | Aspose.Cells worksheet copy example C# | retain conditional formatting when duplicating Excel sheet

using System;
using Aspose.Cells;

// // Loads source.xlsx, copies Sheet1 to a new worksheet while keeping all conditional formatting, renames the copy, and saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        // Load the source workbook
        Workbook workbook = new Workbook("source.xlsx");

        // Get the worksheet to be copied (by name or index)
        Worksheet sourceSheet = workbook.Worksheets["Sheet1"]; // replace with your sheet name

        // Add a new worksheet to hold the copy
        int newSheetIndex = workbook.Worksheets.Add();

        // Copy the source worksheet (including all conditional formatting) to the new worksheet
        sourceSheet.Copy(workbook.Worksheets[newSheetIndex]);

        // Optionally rename the copied worksheet
        workbook.Worksheets[newSheetIndex].Name = "Sheet1_Copy";

        // Save the workbook with the copied worksheet
        workbook.Save("output.xlsx");
    }
}
