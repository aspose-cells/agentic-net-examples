// Title: How to set A4 paper size for every worksheet in an Aspose.Cells workbook using C#
// AI Prompts: Generate C# code that loads an Excel file with Aspose.Cells, iterates over workbook.Worksheets, and assigns PaperSizeType.PaperA4 to each sheet's PageSetup. | Create a C# script that opens a workbook, uses a foreach loop to apply A4 paper size to all worksheets, and saves the updated file with Aspose.Cells.
// Common Searches: C# Aspose.Cells set A4 paper size for all worksheets in a workbook | How to change page setup paper size for every sheet using Aspose.Cells | Loop through worksheets and set print paper size to A4 with Aspose.Cells | Batch update Excel worksheet print settings to A4 in C# | Apply same page layout to multiple worksheets programmatically Aspose.Cells
// Tags: worksheet page setup A4 Aspose.Cells | foreach loop update print settings Aspose.Cells | batch modify Excel paper size C# | apply uniform page layout workbook Aspose.Cells | set paper size for all sheets Aspose.Cells

using System;
using Aspose.Cells;

// Loads an existing workbook, iterates over each worksheet with a foreach loop, sets each sheet's PageSetup.PaperSize to PaperA4, and saves the modified workbook.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Apply A4 paper size to every worksheet using a foreach loop
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Set the paper size to A4
            sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
