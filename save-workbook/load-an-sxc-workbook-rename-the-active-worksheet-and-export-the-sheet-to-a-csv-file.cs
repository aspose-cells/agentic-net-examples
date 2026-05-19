using System;
using Aspose.Cells;

namespace AsposeCellsSxcToCsv
{
    class Program
    {
        static void Main()
        {
            // Path to the source SXC workbook
            string sourcePath = "input.sxc";

            // Load the SXC workbook (uses Workbook(string) constructor)
            Workbook workbook = new Workbook(sourcePath);

            // Get the active worksheet (the one currently selected)
            Worksheet activeSheet = workbook.Worksheets[workbook.Worksheets.ActiveSheetIndex];

            // Rename the active worksheet
            activeSheet.Name = "RenamedSheet";

            // Export the entire workbook (or the active sheet) to CSV format
            // Save method with SaveFormat.Csv follows the provided Save(string, SaveFormat) rule
            string csvPath = "output.csv";
            workbook.Save(csvPath, SaveFormat.Csv);

            Console.WriteLine($"Workbook loaded from '{sourcePath}', sheet renamed, and saved as CSV to '{csvPath}'.");
        }
    }
}