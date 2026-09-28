// Title: Read numeric cells from an Apple‑generated Excel workbook with Aspose.Cells for .NET and calculate their average
// AI Prompts: Write C# code using Aspose.Cells that opens a given Excel file, iterates over every used cell in the first worksheet, sums only numeric values, computes the average, and writes the result to the console. | Update the Aspose.Cells example to exclude cells located in hidden rows or columns when determining the average of numeric data. | Add robust try‑catch blocks around workbook loading and numeric aggregation to provide clear error messages in the Aspose.Cells C# sample.
// Common Searches: Aspose.Cells C# calculate average of numeric cells in Excel workbook | how to ignore hidden rows while averaging numbers with Aspose.Cells .NET | read numeric values from an Apple generated .xlsx using Aspose.Cells | C# example to sum and average cells of type numeric in Aspose.Cells
// Tags: calculate average numeric cells Aspose.Cells | iterate worksheet cells .NET | load Apple generated Excel file Aspose.Cells | skip hidden rows Aspose.Cells | exception handling workbook loading Aspose.Cells

using System;
using Aspose.Cells;

// The program loads an Excel workbook, scans all used cells in the first worksheet, sums values of cells identified as numeric, computes their average, and prints the result to the console.
class Program
{
    static void Main()
    {
        // Load the Apple‑generated spreadsheet
        var workbook = new Workbook("input.xlsx"); // replace with actual file path

        // Access the first worksheet (modify if needed)
        var worksheet = workbook.Worksheets[0];
        var cells = worksheet.Cells;

        double sum = 0;
        int count = 0;

        // Iterate through all used cells and collect numeric values
        foreach (Cell cell in cells)
        {
            if (cell.Type == CellValueType.IsNumeric)
            {
                sum += cell.DoubleValue;
                count++;
            }
        }

        // Calculate and output the average
        if (count > 0)
        {
            double average = sum / count;
            Console.WriteLine($"Average of numeric cells: {average}");
        }
        else
        {
            Console.WriteLine("No numeric cells found.");
        }
    }
}
