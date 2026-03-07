using System;
using Aspose.Cells;

namespace AsposeCellsTimelineExport
{
    public class ExportTimeline
    {
        public static void Main()
        {
            // Path to the source XLSM file that contains the Timeline control
            string sourceFile = "input.xlsm";

            // Desired output path for the XLSX file (Timeline will be retained)
            string destinationFile = "output.xlsx";

            // Load the XLSM workbook
            Workbook workbook = new Workbook(sourceFile);

            // Save the workbook as XLSX
            workbook.Save(destinationFile, SaveFormat.Xlsx);

            Console.WriteLine($"Timeline exported successfully from '{sourceFile}' to '{destinationFile}'.");
        }
    }
}