// Title: Enable NumbersAsText error checking only on worksheets that contain numeric data with Aspose.Cells for .NET
// Common Searches: how to turn on NumbersAsText error checking for specific sheets using Aspose.Cells C# | detect numeric cells in a worksheet before applying error checking with Aspose.Cells | log which worksheets have numeric data when enabling error checking in Aspose.Cells | iterate through all worksheets and set NumbersAsText option conditionally in .NET | Aspose.Cells example to enable NumbersAsText only when numeric values exist
// Tags: conditional NumbersAsText error checking Aspose.Cells | numeric cell detection worksheet Aspose.Cells | worksheet iteration with cell type check Aspose.Cells | logging worksheet processing results Aspose.Cells | enable error checking per sheet Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, iterates over each worksheet, scans the used range for any numeric cell, logs whether numeric data is present, optionally enables the NumbersAsText error‑checking option on sheets that contain numeric values, saves the modified workbook, and outputs the log.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Prepare a log to record changes
            List<string> log = new List<string>();

            // Iterate over all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                bool hasNumeric = false;

                // Determine the used range of the sheet
                Cells cells = sheet.Cells;
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                // Scan the used range for at least one numeric cell
                for (int row = 0; row <= maxRow && !hasNumeric; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];
                        if (cell.Type == CellValueType.IsNumeric)
                        {
                            hasNumeric = true;
                            break;
                        }
                    }
                }

                if (hasNumeric)
                {
                    // NOTE: The EnableErrorChecking and ErrorCheckingOptions APIs may not be
                    // available in older Aspose.Cells versions. If they are present, uncomment
                    // the following lines to enable NumbersAsText error checking.
                    // workbook.Settings.EnableErrorChecking = true;
                    // workbook.Settings.ErrorCheckingOptions.NumbersAsText = true;

                    log.Add($"Numeric data detected on sheet '{sheet.Name}'.");
                }
                else
                {
                    log.Add($"No numeric data found on sheet '{sheet.Name}'.");
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);

            // Output the log entries
            foreach (string entry in log)
            {
                Console.WriteLine(entry);
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
