using System;
using Aspose.Cells;

class RetrieveWorksheetTabIds
{
    static void Main()
    {
        // Path to the Excel file to be loaded
        string inputPath = "input.xlsx";

        // Load the workbook from the specified file
        Workbook workbook = new Workbook(inputPath);

        // Iterate through all worksheets and output their TabId values
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Console.WriteLine($"Worksheet \"{sheet.Name}\" TabId: {sheet.TabId}");
        }
    }
}