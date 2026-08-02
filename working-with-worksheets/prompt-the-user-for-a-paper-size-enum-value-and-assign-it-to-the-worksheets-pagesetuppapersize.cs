// Title: Set Worksheet PageSetup PaperSize from Console Input using Aspose.Cells for .NET (C#)
// Description: Shows how to read a paper size name from the console, convert it to the Aspose.Cells PaperSizeType enum (case‑insensitive), assign it to Worksheet.PageSetup.PaperSize, and save the workbook as PaperSizeDemo.xlsx, with graceful handling of invalid entries.
// Keywords: Aspose.Cells | C# PaperSize | Worksheet PageSetup | PaperSizeType enum | set Excel paper size programmatically | console input | Aspose.Cells .NET | Excel export paper size | PageSetup.PaperSize | PaperA4 | PaperLetter | PaperLegal
// Common Searches: Aspose.Cells change worksheet paper size from console | C# parse string to PaperSizeType enum | Set Excel sheet paper size programmatically Aspose.Cells | PageSetup PaperSize example C# | How to assign custom paper size in Aspose.Cells workbook
// Developer Intent: Apply a user‑provided PaperSizeType value to a worksheet’s PageSetup.PaperSize.
// Use Cases: Console utility that lets end‑users select A4, Letter, or Legal before generating an Excel report. | Dynamic paper‑size adjustment when exporting a workbook to PDF or XPS. | Validation of user‑entered paper size strings with fallback to the default size.
// AI Prompts: Generate C# code that prompts for a paper size, safely converts it to PaperSizeType, sets Worksheet.PageSetup.PaperSize, and saves the workbook. | Show how to handle an invalid paper size entry by keeping the default size and displaying an error message in Aspose.Cells. | Create an example that iterates over a list of paper sizes, creates separate worksheets, and assigns each the corresponding PageSetup.PaperSize.

using System;
using Aspose.Cells;

namespace AsposeCellsPaperSizeDemo
{
    // Shows how to read a paper size name from the console, convert it to the Aspose.Cells PaperSizeType enum (case‑insensitive), assign it to Worksheet.PageSetup.PaperSize, and save the workbook as PaperSizeDemo.xlsx, with graceful handling of invalid entries.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Prompt the user to enter a paper size enum name (e.g., PaperA4, PaperLetter)
            Console.WriteLine("Enter a paper size (e.g., PaperA4, PaperLetter, PaperLegal):");
            string input = Console.ReadLine();

            // Try to parse the user input to the PaperSizeType enum (ignore case)
            if (Enum.TryParse<PaperSizeType>(input, true, out PaperSizeType paperSize))
            {
                // Assign the selected paper size to the worksheet's PageSetup
                worksheet.PageSetup.PaperSize = paperSize;
                Console.WriteLine($"Paper size set to: {worksheet.PageSetup.PaperSize}");
            }
            else
            {
                Console.WriteLine("Invalid paper size entered. Using default size.");
            }

            // Save the workbook to a file
            workbook.Save("PaperSizeDemo.xlsx");
            Console.WriteLine("Workbook saved as PaperSizeDemo.xlsx");
        }
    }
}
