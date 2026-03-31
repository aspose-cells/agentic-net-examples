using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeCopyDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load the source workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (you can change the index or name as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Define the source range to copy (e.g., cells A1:C5)
            AsposeRange sourceRange = sheet.Cells.CreateRange("A1:C5");

            // Define the destination range where the data will be pasted (e.g., starting at E1)
            AsposeRange destinationRange = sheet.Cells.CreateRange("E1:G5");

            // Configure paste options
            PasteOptions pasteOptions = new PasteOptions
            {
                PasteType = PasteType.All,
                SkipBlanks = false,
                Transpose = false,
                OperationType = PasteOperationType.None,
                IgnoreLinksToOriginalFile = true
            };

            // Perform the copy with the specified options
            destinationRange.Copy(sourceRange, pasteOptions);

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("output.xlsx");
        }
    }
}