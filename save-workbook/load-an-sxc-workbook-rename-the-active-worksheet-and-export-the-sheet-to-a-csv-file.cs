using System;
using Aspose.Cells;

namespace AsposeCellsSxcToCsv
{
    class Program
    {
        static void Main()
        {
            // Path to the source SXC file
            string sourcePath = "input.sxc";

            // Load the SXC workbook using the string constructor
            Workbook workbook = new Workbook(sourcePath);

            // Get the active worksheet (the one currently selected)
            Worksheet activeSheet = workbook.Worksheets[workbook.Worksheets.ActiveSheetIndex];

            // Rename the active worksheet
            activeSheet.Name = "RenamedSheet";

            // Export the workbook (first worksheet) to CSV format
            // Save method with file name and SaveFormat enum is used as per the provided rules
            workbook.Save("output.csv", SaveFormat.Csv);

            // Optional: inform the user
            Console.WriteLine("Workbook loaded, worksheet renamed, and exported to CSV successfully.");
        }
    }
}