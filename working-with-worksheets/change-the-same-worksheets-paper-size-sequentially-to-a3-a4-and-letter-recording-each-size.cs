// Title: Aspose.Cells C# – Change a Worksheet’s Paper Size to A3, A4, then Letter and Log the Sequence
// Description: Creates a new Workbook, accesses the first Worksheet, sets PageSetup.PaperSize to A3, A4, and Letter in order, records each size in a List, outputs the log to the console, and saves the file as PaperSizeSequence.xlsx.
// Keywords: Aspose.Cells C# | worksheet paper size | PageSetup PaperSize | PaperSizeType A3 | PaperSizeType A4 | PaperSizeType Letter | record paper size sequence | save workbook after page setup | C# Excel automation
// Common Searches: Aspose.Cells change worksheet paper size to A3 | set worksheet paper size to A4 using C# | how to log paper size changes with Aspose.Cells | C# example for sequential PaperSizeType values | save Excel file after modifying page setup Aspose
// Developer Intent: Set a worksheet’s PaperSize to A3, then A4, then Letter while capturing each value in a collection.
// Use Cases: Generate a multi‑section report where each section requires a different standard paper size. | Automate printer‑setting validation by cycling through common formats and storing the results for quality checks. | Create an audit trail of page‑setup changes during workbook generation for compliance or debugging.
// AI Prompts: Show how to reset the worksheet’s paper size to the default after logging the sequence. | Provide code to export the recorded PaperSizeType values to a CSV file with Aspose.Cells. | Explain how to apply the A3‑A4‑Letter paper size sequence to every worksheet in a workbook.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsPaperSizeDemo
{
    // Creates a new Workbook, accesses the first Worksheet, sets PageSetup.PaperSize to A3, A4, and Letter in order, records each size in a List, outputs the log to the console, and saves the file as PaperSizeSequence.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (using the standard Aspose.Cells creation rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // List to record the paper sizes applied
            List<PaperSizeType> recordedSizes = new List<PaperSizeType>();

            // 1. Set paper size to A3 and record
            sheet.PageSetup.PaperSize = PaperSizeType.PaperA3;
            recordedSizes.Add(sheet.PageSetup.PaperSize);
            Console.WriteLine("Paper size set to: " + sheet.PageSetup.PaperSize);

            // 2. Set paper size to A4 and record
            sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;
            recordedSizes.Add(sheet.PageSetup.PaperSize);
            Console.WriteLine("Paper size set to: " + sheet.PageSetup.PaperSize);

            // 3. Set paper size to Letter and record
            sheet.PageSetup.PaperSize = PaperSizeType.PaperLetter;
            recordedSizes.Add(sheet.PageSetup.PaperSize);
            Console.WriteLine("Paper size set to: " + sheet.PageSetup.PaperSize);

            // Output the recorded sequence
            Console.WriteLine("\nRecorded paper sizes in order:");
            foreach (var size in recordedSizes)
            {
                Console.WriteLine("- " + size);
            }

            // Save the workbook (using the standard Aspose.Cells save rule)
            workbook.Save("PaperSizeSequence.xlsx");
        }
    }
}
