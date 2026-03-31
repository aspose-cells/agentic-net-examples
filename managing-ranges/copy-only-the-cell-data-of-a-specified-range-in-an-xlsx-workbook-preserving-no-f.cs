using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class Program
    {
        static void Main()
        {
            // Load the source workbook (replace with your actual file path)
            Workbook workbook = new Workbook("source.xlsx");

            // Get the first worksheet as the source sheet
            Worksheet sourceSheet = workbook.Worksheets[0];
            sourceSheet.Name = "Source";

            // Add a new worksheet that will receive the copied data
            int destIndex = workbook.Worksheets.Add();
            Worksheet destinationSheet = workbook.Worksheets[destIndex];
            destinationSheet.Name = "Destination";

            // Define the source range you want to copy (e.g., A1:B5)
            Aspose.Cells.Range sourceRange = sourceSheet.Cells.CreateRange("A1:B5");

            // Define the destination range where the data will be placed (e.g., C1:D5)
            Aspose.Cells.Range destinationRange = destinationSheet.Cells.CreateRange("C1:D5");

            // Copy only the cell data (values and formulas) without any formatting or metadata
            destinationRange.CopyData(sourceRange);

            // Save the workbook with the copied data
            workbook.Save("output.xlsx");
        }
    }
}