using System;
using Aspose.Cells;

class LoadCsvWithFormulas
{
    static void Main()
    {
        // Path to the CSV file that contains formulas (e.g., "=SUM(B1:B3)")
        string csvPath = "input.csv";

        // Create an empty workbook (lifecycle start)
        Workbook workbook = new Workbook();

        // Get the Cells collection of the first worksheet
        Cells cells = workbook.Worksheets[0].Cells;

        // Configure TxtLoadOptions:
        // - Separator set to ',' for CSV
        // - HasFormula = true so strings starting with '=' are treated as formulas
        // - ConvertNumericData = true to convert numeric strings to numbers
        TxtLoadOptions txtOptions = new TxtLoadOptions
        {
            Separator = ',',
            HasFormula = true,
            ConvertNumericData = true
        };

        // Import the CSV data starting at cell A1 (row 0, column 0)
        cells.ImportCSV(csvPath, txtOptions, 0, 0);

        // Parse and calculate all imported formulas
        workbook.CalculateFormula();

        // Example: output the evaluated value of cell A1 (adjust as needed)
        Console.WriteLine("A1 evaluated value: " + cells["A1"].Value);

        // Save the workbook to an Excel file (lifecycle end)
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}