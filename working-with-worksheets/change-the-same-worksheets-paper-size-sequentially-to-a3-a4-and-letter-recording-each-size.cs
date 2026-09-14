// Title: Changing a worksheet’s page setup paper size to A3, A4, and Letter sequentially using Aspose.Cells for .NET
// AI Prompts: Write C# code that sets Worksheet.PageSetup.PaperSize to PaperA3, then PaperA4, then PaperLetter, and stores each size name in a list. | Show how to output the recorded paper sizes to the console after applying multiple page size settings with Aspose.Cells.
// Common Searches: Aspose.Cells set worksheet paper size to A3 then A4 then Letter in C# | How to log each paper size change when modifying PageSetup in Aspose.Cells | C# example for sequentially updating Excel worksheet page setup paper sizes with Aspose.Cells
// Tags: Aspose.Cells worksheet page setup paper size | C# set PaperSizeType A3 A4 Letter | record paper size changes Aspose.Cells | sequential page size update Excel .NET | save workbook after custom paper sizes Aspose

using Aspose.Cells;
using System;
using System.Collections.Generic;

// The sample creates a new workbook, accesses the first worksheet, sequentially sets its PageSetup.PaperSize to A3, A4, and Letter, records each applied size in a list, prints the recorded sizes, and saves the file as PaperSizeDemo.xlsx.
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

        // Apply A3 paper size and record it
        sheet.PageSetup.PaperSize = PaperSizeType.PaperA3;
        recordedSizes.Add("A3");

        // Apply A4 paper size and record it
        sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;
        recordedSizes.Add("A4");

        // Apply Letter paper size and record it
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
