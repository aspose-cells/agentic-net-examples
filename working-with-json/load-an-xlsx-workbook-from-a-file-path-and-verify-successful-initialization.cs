using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the Excel file to be loaded
        string filePath = "sample.xlsx";

        // Load the workbook using the string constructor (provided rule)
        Workbook workbook = new Workbook(filePath);

        // Verify successful initialization by checking the worksheet collection
        if (workbook.Worksheets != null && workbook.Worksheets.Count > 0)
        {
            Console.WriteLine($"Workbook loaded successfully. Worksheet count: {workbook.Worksheets.Count}");
        }
        else
        {
            Console.WriteLine("Workbook loaded, but no worksheets were found.");
        }
    }
}