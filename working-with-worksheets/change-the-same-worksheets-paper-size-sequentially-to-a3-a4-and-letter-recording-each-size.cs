using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace PaperSizeSequenceDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (using the standard creation rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // List to record the paper sizes applied
            List<PaperSizeType> appliedSizes = new List<PaperSizeType>();

            // Set paper size to A3 and record
            worksheet.PageSetup.PaperSize = PaperSizeType.PaperA3;
            appliedSizes.Add(worksheet.PageSetup.PaperSize);

            // Set paper size to A4 and record
            worksheet.PageSetup.PaperSize = PaperSizeType.PaperA4;
            appliedSizes.Add(worksheet.PageSetup.PaperSize);

            // Set paper size to Letter and record
            worksheet.PageSetup.PaperSize = PaperSizeType.PaperLetter;
            appliedSizes.Add(worksheet.PageSetup.PaperSize);

            // Output the recorded paper sizes to the console
            Console.WriteLine("Applied paper sizes in sequence:");
            foreach (PaperSizeType size in appliedSizes)
            {
                Console.WriteLine("- " + size);
            }

            // Save the workbook (using the standard save rule)
            workbook.Save("PaperSizeSequence.xlsx");
        }
    }
}