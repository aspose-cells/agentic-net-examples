// Title: Set black‑and‑white printing for all worksheets in an Excel file using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set black and white print mode for every sheet | How to reduce ink usage by enabling monochrome printing in a workbook with Aspose.Cells | C# code to apply PageSetup.BlackAndWhite to all worksheets in an XLSX file | Bulk configure print settings across worksheets using Aspose.Cells .NET | Programmatically turn on black‑and‑white printing for an entire workbook in C#
// Tags: Aspose.Cells BlackAndWhite page setup | set monochrome printing for workbook | bulk worksheet print configuration .NET | ink saving print settings Excel | C# PageSetup BlackAndWhite usage

using Aspose.Cells;
using System;

// Loads an existing XLSX file, iterates through each worksheet, sets the PageSetup.BlackAndWhite property to true to enable monochrome printing, and saves the workbook.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Enable black‑and‑white printing for every worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.PageSetup.BlackAndWhite = true;
        }

        // Save the workbook with the updated printing settings
        workbook.Save("output.xlsx");
    }
}
