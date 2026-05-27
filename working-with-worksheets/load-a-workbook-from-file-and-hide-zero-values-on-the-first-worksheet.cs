using System;
using Aspose.Cells;

namespace AsposeCellsHideZeroValues
{
    class Program
    {
        static void Main()
        {
            // Path to the existing Excel file
            string inputPath = "input.xlsx";

            // Load the workbook from the file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (index 0)
            Worksheet firstSheet = workbook.Worksheets[0];

            // Hide zero values on this worksheet
            firstSheet.DisplayZeros = false;

            // Save the modified workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
        }
    }
}