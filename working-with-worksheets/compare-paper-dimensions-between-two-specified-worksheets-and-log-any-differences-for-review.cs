using System;
using Aspose.Cells;

namespace PaperDimensionComparer
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file containing the worksheets to compare
            string inputFile = "InputWorkbook.xlsx";

            // Names (or indexes) of the two worksheets to compare
            string firstSheetName = "Sheet1";
            string secondSheetName = "Sheet2";

            // Load the workbook (creation and loading follow Aspose.Cells standard pattern)
            Workbook workbook = new Workbook(inputFile);

            // Retrieve the two worksheets
            Worksheet ws1 = workbook.Worksheets[firstSheetName];
            Worksheet ws2 = workbook.Worksheets[secondSheetName];

            if (ws1 == null || ws2 == null)
            {
                Console.WriteLine("One or both specified worksheets were not found.");
                return;
            }

            // Access PageSetup for each worksheet
            PageSetup ps1 = ws1.PageSetup;
            PageSetup ps2 = ws2.PageSetup;

            // Compare paper width
            double width1 = ps1.PaperWidth;   // inches, considers orientation
            double width2 = ps2.PaperWidth;

            // Compare paper height
            double height1 = ps1.PaperHeight;
            double height2 = ps2.PaperHeight;

            // Compare paper size enum (useful when custom sizes are not set)
            PaperSizeType size1 = ps1.PaperSize;
            PaperSizeType size2 = ps2.PaperSize;

            bool differencesFound = false;

            // Log differences in width
            if (Math.Abs(width1 - width2) > 0.0001) // tolerance for floating point comparison
            {
                Console.WriteLine($"Width differs: '{firstSheetName}' = {width1} inches, '{secondSheetName}' = {width2} inches");
                differencesFound = true;
            }

            // Log differences in height
            if (Math.Abs(height1 - height2) > 0.0001)
            {
                Console.WriteLine($"Height differs: '{firstSheetName}' = {height1} inches, '{secondSheetName}' = {height2} inches");
                differencesFound = true;
            }

            // Log differences in PaperSize enum (covers standard sizes)
            if (size1 != size2)
            {
                Console.WriteLine($"PaperSize enum differs: '{firstSheetName}' = {size1}, '{secondSheetName}' = {size2}");
                differencesFound = true;
            }

            if (!differencesFound)
            {
                Console.WriteLine("No differences in paper dimensions were found between the two worksheets.");
            }

            // Optionally, save the workbook if any modifications were made (here we just demonstrate saving)
            string outputFile = "OutputWorkbook.xlsx";
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved to '{outputFile}'.");
        }
    }
}