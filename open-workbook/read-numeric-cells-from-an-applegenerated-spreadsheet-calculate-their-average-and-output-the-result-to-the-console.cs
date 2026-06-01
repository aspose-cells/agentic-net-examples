using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the Apple‑generated spreadsheet
        string filePath = "input.xlsx";

        // Load the workbook (load rule)
        Workbook workbook = new Workbook(filePath);

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        double sum = 0;
        int count = 0;

        // Iterate through all cells in the worksheet
        foreach (Cell cell in cells)
        {
            // Consider only numeric cells (including dates, which are stored as numbers)
            if (cell.IsNumericValue)
            {
                sum += cell.DoubleValue;
                count++;
            }
        }

        if (count > 0)
        {
            double average = sum / count;
            Console.WriteLine($"Average of numeric cells: {average}");
        }
        else
        {
            Console.WriteLine("No numeric cells found in the worksheet.");
        }
    }
}