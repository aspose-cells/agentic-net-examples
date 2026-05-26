using System;
using System.IO;
using Aspose.Cells;

class FormulaErrorLogger
{
    static void Main()
    {
        // Paths for input workbook, output log, and (optional) saved workbook
        string inputPath = "input.xlsx";
        string logPath = "FormulaErrorLog.txt";
        string outputPath = "output.xlsx";

        // Load the workbook
        Workbook workbook = new Workbook(inputPath);

        // Calculate all formulas while ignoring errors so processing can continue
        workbook.CalculateFormula(new CalculationOptions { IgnoreError = true });

        // Open a writer for the log file
        using (StreamWriter writer = new StreamWriter(logPath))
        {
            // Write CSV header
            writer.WriteLine("CellAddress,ErrorType,Timestamp");

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                // Scan all used cells
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];

                        // Process only cells that contain a formula
                        if (cell.IsFormula)
                        {
                            // Retrieve rich value to inspect possible error information
                            CellRichValue richValue = cell.GetRichValue();

                            // If the cell has an error, log details
                            if (richValue != null && richValue.ErrorValue != 0)
                            {
                                string address = cell.Name; // e.g., "A1"
                                string errorType = richValue.ErrorValue.ToString(); // enum name
                                string timestamp = DateTime.Now.ToString("o"); // ISO 8601 format

                                writer.WriteLine($"{address},{errorType},{timestamp}");
                            }
                        }
                    }
                }
            }
        }

        // Save the workbook (unchanged except for calculated values)
        workbook.Save(outputPath);
    }
}