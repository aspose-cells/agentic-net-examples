// Title: C# console program using Aspose.Cells to display each worksheet’s paper size, width and height
// Common Searches: aspocells c# get worksheet page setup paper size and dimensions | list paper width and height of all sheets in an Excel file using Aspose.Cells | how to read paper dimensions of worksheets in a workbook with Aspose.Cells C# | console app to output Excel sheet paper size properties via Aspose.Cells | retrieve page setup properties for each worksheet using Aspose.Cells in .NET
// Tags: Aspose.Cells worksheet page setup extraction | C# read Excel paper dimensions | Aspose.Cells get paper size points | console application workbook sheet properties | iterate worksheets page setup Aspose

using System;
using Aspose.Cells;

namespace AsposeCellsPaperDimensions
{
    // The program loads an Excel workbook (default "input.xlsx" or a path supplied via command line), iterates through all worksheets, and writes each sheet's name along with its PageSetup PaperSize, PaperWidth, and PaperHeight (in points) to the console before exiting.
    class Program
    {
        static void Main(string[] args)
        {
            // Determine the workbook file path.
            // If a path is provided as a command‑line argument it is used,
            // otherwise a default file name is assumed.
            string workbookPath = args.Length > 0 ? args[0] : "input.xlsx";

            // Load the workbook from the specified file.
            Workbook workbook = new Workbook(workbookPath);

            // Iterate through each worksheet in the workbook.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Access the PageSetup object which contains paper settings.
                PageSetup pageSetup = sheet.PageSetup;

                // Output worksheet name and its paper dimensions.
                Console.WriteLine($"Worksheet: {sheet.Name}");
                Console.WriteLine($"Paper Size   : {pageSetup.PaperSize}");
                Console.WriteLine($"Paper Width  : {pageSetup.PaperWidth} points");
                Console.WriteLine($"Paper Height : {pageSetup.PaperHeight} points");
                Console.WriteLine(); // Blank line for readability
            }

            // Application ends after printing the information.
        }
    }
}
