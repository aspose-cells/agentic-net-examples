// Title: Set Worksheet Paper Size Sequentially to A3, A4, and Letter with Aspose.Cells for .NET
// Description: This example creates a new Workbook, accesses the first Worksheet, changes its PageSetup.PaperSize to A3, then A4, then Letter, records each size in a list, prints the sequence to the console, and saves the file as PaperSizeSequence.xlsx.
// Keywords: Aspose.Cells C# set paper size | Worksheet PageSetup PaperSizeType | change worksheet paper size A3 A4 Letter | record paper size changes Aspose.Cells | save workbook after page setup update
// Common Searches: Aspose.Cells change worksheet paper size to A3 then A4 then Letter | how to record paper size after setting it in Aspose.Cells | C# Aspose.Cells multiple paper size settings on one sheet | retrieve current PaperSizeType value in Aspose.Cells
// Developer Intent: Set the worksheet’s paper size to A3, then A4, then Letter while capturing each size in a collection.
// Use Cases: Generate a multi‑section report where each section must be printed on a different paper format and log the applied format for audit. | Validate that PageSetup changes take effect by logging the PaperSize after each assignment during automated UI testing. | Create a test matrix that cycles through common paper sizes to verify print preview rendering in a CI pipeline.
// AI Prompts: Write C# code that changes a worksheet’s paper size to A5, stores the previous size, and restores it after exporting to PDF using Aspose.Cells. | Explain how to map each PaperSizeType enum value to its physical dimensions in inches and millimeters. | Show how to serialize a List<PaperSizeType> to JSON and write it to a file for later analysis.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace PaperSizeSequenceDemo
{
    // This example creates a new Workbook, accesses the first Worksheet, changes its PageSetup.PaperSize to A3, then A4, then Letter, records each size in a list, prints the sequence to the console, and saves the file as PaperSizeSequence.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the PageSetup object for the worksheet
            PageSetup pageSetup = worksheet.PageSetup;

            // List to record the paper sizes after each change
            List<PaperSizeType> recordedSizes = new List<PaperSizeType>();

            // Change to A3 and record
            pageSetup.PaperSize = PaperSizeType.PaperA3;
            recordedSizes.Add(pageSetup.PaperSize);
            Console.WriteLine("Set paper size to A3");

            // Change to A4 and record
            pageSetup.PaperSize = PaperSizeType.PaperA4;
            recordedSizes.Add(pageSetup.PaperSize);
            Console.WriteLine("Set paper size to A4");

            // Change to Letter and record
            pageSetup.PaperSize = PaperSizeType.PaperLetter;
            recordedSizes.Add(pageSetup.PaperSize);
            Console.WriteLine("Set paper size to Letter");

            // Output the recorded paper sizes
            Console.WriteLine("\nRecorded paper sizes in sequence:");
            foreach (PaperSizeType size in recordedSizes)
            {
                Console.WriteLine("- " + size);
            }

            // Save the workbook (lifecycle save)
            workbook.Save("PaperSizeSequence.xlsx");
        }
    }
}
