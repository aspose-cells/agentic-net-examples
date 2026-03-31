using System;
using Aspose.Cells;

namespace WorksheetTransferExample
{
    public class Program
    {
        public static void Main()
        {
            // Load the source workbook (replace the path with your actual file)
            Workbook sourceWorkbook = new Workbook("source.xlsx");

            // Get the worksheet you want to transfer (here we use the first worksheet)
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

            // Create a new (empty) destination workbook
            Workbook destWorkbook = new Workbook();

            // Get the first worksheet in the destination workbook
            Worksheet destSheet = destWorkbook.Worksheets[0];

            // Copy all cell data, formatting, and formulas from the source sheet to the destination sheet
            destSheet.Copy(sourceSheet);

            // Save the destination workbook
            destWorkbook.Save("destination.xlsx");
        }
    }
}