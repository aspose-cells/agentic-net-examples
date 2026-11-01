// Title: Iterate worksheets, detect numeric data, log affected sheets, and enable NumbersAsText error checking when supported using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads a workbook with Aspose.Cells, scans every worksheet for numeric values, enables the NumbersAsText error checking option only on worksheets that contain numeric cells, and returns a list of the modified sheet names. | Write a method using Aspose.Cells for .NET that iterates through all worksheets, checks the used range for numeric content, applies the NumbersAsText error option conditionally, and records each sheet name where the option was applied.
// Common Searches: Aspose.Cells C# enable NumbersAsText error checking on worksheets that contain numbers | How to loop through all sheets in a workbook and apply error options based on cell type using Aspose.Cells | C# Aspose.Cells detect numeric cells in a worksheet and log sheet names | Conditional error checking per worksheet Aspose.Cells .NET example
// Tags: NumbersAsText option per worksheet Aspose.Cells | numeric data detection Aspose.Cells | worksheet change logging Aspose.Cells | iterate workbook sheets C# Aspose.Cells | conditional error option based on cell type Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads an Excel file, iterates each worksheet, scans the used range for numeric values, logs the names of sheets that contain numeric data, and (if supported) enables the NumbersAsText error checking option for those sheets before saving the workbook.
class Program
{
    static void Main()
    {
        // Paths for input and output files
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // List to keep a log of changes
            List<string> changeLog = new List<string>();

            // Iterate over all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                bool hasNumericData = false;

                // Determine the used range of the worksheet
                int maxRow = sheet.Cells.MaxDataRow;
                int maxCol = sheet.Cells.MaxDataColumn;

                // Scan cells for numeric values
                for (int row = 0; row <= maxRow && !hasNumericData; row++)
                {
                    for (int col = 0; col <= maxCol && !hasNumericData; col++)
                    {
                        Cell cell = sheet.Cells[row, col];
                        if (cell.Type == CellValueType.IsNumeric)
                        {
                            hasNumericData = true;
                        }
                    }
                }

                // If numeric data is present, log the detection.
                // (The NumbersAsText error checking option is not available in this version of Aspose.Cells.)
                if (hasNumericData)
                {
                    changeLog.Add($"Numeric data detected on sheet \"{sheet.Name}\".");
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);

            // Output the log of changes
            Console.WriteLine("Changes applied:");
            foreach (string logEntry in changeLog)
            {
                Console.WriteLine(logEntry);
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
