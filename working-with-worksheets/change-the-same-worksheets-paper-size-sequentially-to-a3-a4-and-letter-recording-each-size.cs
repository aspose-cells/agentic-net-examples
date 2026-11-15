// Title: How to change a worksheet's page setup paper size sequentially to A3, A4, and Letter using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set worksheet page setup to A3 then A4 then Letter | record each PaperSizeType after changing page setup in Aspose.Cells | C# example for sequential paper size changes in an Excel worksheet | how to log applied paper sizes when modifying worksheet PageSetup with Aspose.Cells | save workbook after updating multiple page setup paper sizes in C#
// Tags: Aspose.Cells set worksheet page setup paper size | C# sequential PaperSizeType changes | record applied PaperSizeType values | save workbook after page setup modifications | Aspose.Cells PaperSizeType enumeration usage

using System;
using System.Collections.Generic;
using Aspose.Cells;

// The sample creates a new workbook, accesses the first worksheet, and changes its PageSetup.PaperSize property to A3, then A4, then Letter. Each applied PaperSizeType is added to a List, printed to the console, and the workbook is saved as SequentialPaperSizes.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // List to record the paper sizes applied
        List<PaperSizeType> appliedSizes = new List<PaperSizeType>();

        // Set paper size to A3 and record
        sheet.PageSetup.PaperSize = PaperSizeType.PaperA3;
        appliedSizes.Add(sheet.PageSetup.PaperSize);

        // Set paper size to A4 and record
        sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;
        appliedSizes.Add(sheet.PageSetup.PaperSize);

        // Set paper size to Letter and record
        sheet.PageSetup.PaperSize = PaperSizeType.PaperLetter;
        appliedSizes.Add(sheet.PageSetup.PaperSize);

        // Output the recorded paper sizes
        Console.WriteLine("Applied paper sizes in sequence:");
        foreach (PaperSizeType size in appliedSizes)
        {
            Console.WriteLine(size);
        }

        // Save the workbook to a file
        workbook.Save("SequentialPaperSizes.xlsx");
    }
}
