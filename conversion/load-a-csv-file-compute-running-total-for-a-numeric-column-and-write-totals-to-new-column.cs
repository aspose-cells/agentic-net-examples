using System;
using Aspose.Cells;

class RunningTotalCsvProcessor
{
    static void Main()
    {
        // Paths for input CSV and output Excel file
        string csvPath = "input.csv";
        string outputPath = "output.xlsx";

        // Create a new workbook (lifecycle rule)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Import CSV data starting at cell A1 (row 0, column 0)
        // Using comma as delimiter, converting numeric data automatically
        cells.ImportCSV(csvPath, ",", true, 0, 0); // ImportCSV rule

        // Assume the first row contains headers and the numeric column is the second column (index 1)
        int headerRow = 0;
        int numericColumnIndex = 1;

        // Determine the last row with data
        int lastDataRow = cells.MaxDataRow;

        // Add a new column for the running total after the existing data columns
        int totalColumnIndex = cells.MaxDataColumn + 1;
        cells[headerRow, totalColumnIndex].PutValue("RunningTotal");

        // Compute running total and write to the new column
        double runningSum = 0;
        for (int row = headerRow + 1; row <= lastDataRow; row++)
        {
            // Get the numeric value from the specified column; if not numeric, treat as 0
            double currentValue = cells[row, numericColumnIndex].DoubleValue;
            runningSum += currentValue;
            cells[row, totalColumnIndex].PutValue(runningSum);
        }

        // Save the workbook (lifecycle rule)
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}