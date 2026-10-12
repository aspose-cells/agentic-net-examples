// Title: Change a worksheet’s page setup paper size to A3, A4, and Letter sequentially with Aspose.Cells for .NET
// AI Prompts: Assign PaperSizeType.PaperA3 to sheet.PageSetup.PaperSize, then PaperSizeType.PaperA4, then PaperSizeType.PaperLetter, adding each size label to a List<string>. | Iterate over the List<string> to write each recorded size to the console and finally save the workbook as PaperSizeDemo.xlsx.
// Common Searches: asp.net set worksheet paper size to A3 using Aspose.Cells | c# change Aspose.Cells worksheet page setup to A4 then Letter | how to log paper size changes in an Aspose.Cells workbook | using PaperSizeType enumeration to modify page setup in Aspose.Cells C# | save workbook after updating page setup paper size Aspose.Cells
// Tags: Aspose.Cells worksheet page setup paper size | C# PaperSizeType enumeration usage | record applied paper sizes in Aspose.Cells | save workbook after page setup modification Aspose.Cells | sequential paper size changes Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections.Generic;

// // Demonstrates creating a workbook, accessing the first worksheet, sequentially setting its PageSetup.PaperSize to A3, A4, and Letter, recording each size in a list, printing the recorded sizes, and saving the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // List to record the paper sizes applied
        List<string> recordedSizes = new List<string>();

        // Change paper size to A3 and record
        sheet.PageSetup.PaperSize = PaperSizeType.PaperA3;
        recordedSizes.Add("A3");

        // Change paper size to A4 and record
        sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;
        recordedSizes.Add("A4");

        // Change paper size to Letter and record
        sheet.PageSetup.PaperSize = PaperSizeType.PaperLetter;
        recordedSizes.Add("Letter");

        // Output the recorded paper sizes
        foreach (string size in recordedSizes)
        {
            Console.WriteLine($"Recorded paper size: {size}");
        }

        // Save the workbook (optional)
        workbook.Save("PaperSizeDemo.xlsx");
    }
}
