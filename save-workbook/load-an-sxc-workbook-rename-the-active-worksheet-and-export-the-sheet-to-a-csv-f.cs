using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source SXC workbook
        string sourcePath = "input.sxc";

        // Load the workbook from the SXC file
        Workbook workbook = new Workbook(sourcePath);

        // Rename the active worksheet (the first worksheet in the collection)
        Worksheet activeSheet = workbook.Worksheets[0];
        activeSheet.Name = "RenamedSheet";

        // Export the active worksheet to a CSV file
        string csvPath = "output.csv";
        workbook.Save(csvPath, SaveFormat.Csv);
    }
}