using System;
using Aspose.Cells;

namespace AsposeCellsChartCopyDemo
{
    class Program
    {
        static void Main()
        {
            // Load the source workbook that contains the chart sheet.
            // Replace "source.xlsx" with the actual path to your source file.
            Workbook sourceWorkbook = new Workbook("source.xlsx");

            // Assume the chart is on the first worksheet (index 0).
            // If you know the name, you can use sourceWorkbook.Worksheets["ChartSheetName"].
            Worksheet sourceChartSheet = sourceWorkbook.Worksheets[0];

            // Create a new (empty) destination workbook.
            Workbook destWorkbook = new Workbook();

            // Add a copy of the chart sheet to the destination workbook.
            // The AddCopy method copies the entire worksheet, including charts,
            // shapes, and all formatting.
            int copiedIndex = destWorkbook.Worksheets.AddCopy(sourceChartSheet.Index);

            // Optionally rename the copied chart sheet.
            Worksheet copiedChartSheet = destWorkbook.Worksheets[copiedIndex];
            copiedChartSheet.Name = "CopiedChartSheet";

            // Save the destination workbook.
            // Replace "output.xlsx" with the desired output path.
            destWorkbook.Save("output.xlsx");
        }
    }
}