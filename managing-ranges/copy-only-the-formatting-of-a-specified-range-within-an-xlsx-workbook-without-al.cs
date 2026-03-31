using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    class Program
    {
        static void Main(string[] args)
        {
            CopyRangeFormattingOnly.Run();
        }
    }

    public class CopyRangeFormattingOnly
    {
        public static void Run()
        {
            // Load the workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (adjust index or name as needed)
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Define the source range whose formatting will be copied (e.g., A1:C5)
            AsposeRange sourceRange = cells.CreateRange("A1:C5");

            // Define the destination range where the formatting will be applied (e.g., E1:G5)
            AsposeRange destinationRange = cells.CreateRange("E1:G5");

            // Copy only the style (formatting) from source to destination
            destinationRange.CopyStyle(sourceRange);

            // Save the workbook with the updated formatting
            workbook.Save("output.xlsx");
        }
    }
}