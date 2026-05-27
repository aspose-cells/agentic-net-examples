using System;
using Aspose.Cells;

namespace WorksheetCopyExample
{
    class Program
    {
        static void Main()
        {
            // Load the source workbook (replace with your actual file path)
            Workbook sourceWorkbook = new Workbook("source.xlsx");

            // Get the worksheet you want to copy (first worksheet in this case)
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

            // Create a new (empty) destination workbook
            Workbook destWorkbook = new Workbook();

            // Get the destination worksheet (first worksheet by default)
            Worksheet destSheet = destWorkbook.Worksheets[0];

            // Copy the source worksheet to the destination worksheet.
            // This method copies contents, formats, and preserves formulas.
            destSheet.Copy(sourceSheet);

            // Save the destination workbook (replace with your desired output path)
            destWorkbook.Save("destination.xlsx");
        }
    }
}