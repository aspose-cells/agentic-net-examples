// Title: How to set A4 paper size for every worksheet in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel file with Aspose.Cells, iterates through all worksheets, and sets each sheet's PageSetup.PaperSize to PaperSizeType.PaperA4. | Create a reusable C# method that accepts a Workbook object and applies A4 paper size to every worksheet using a foreach loop. | Generate a C# example that changes both the paper size to A4 and the orientation to portrait for all worksheets in a workbook with Aspose.Cells.
// Common Searches: Aspose.Cells C# set A4 paper size for all worksheets in a workbook | foreach loop change page setup paper size Aspose.Cells .NET | apply same page settings to every sheet using Aspose.Cells | batch update worksheet page setup to A4 with Aspose.Cells C#
// Tags: Aspose.Cells set worksheet paper size | C# foreach apply page setup to Excel sheets | A4 paper size configuration Aspose.Cells | batch update page settings .NET workbook | apply default page layout to all worksheets

using Aspose.Cells;
using System;

// Loads an existing workbook, loops through each worksheet, sets PageSetup.PaperSize to PaperA4, and saves the updated file.
class Program
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Apply A4 paper size to every worksheet using a foreach loop
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Set the paper size to A4
            sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
