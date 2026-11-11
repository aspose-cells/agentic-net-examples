// Title: Load an Excel workbook and process only visible worksheets in C# with Aspose.Cells
// Common Searches: how to load an Excel file and iterate only visible worksheets with Aspose.Cells C# | skip hidden worksheets during Excel processing using Aspose.Cells .NET | trim whitespace from cell values while looping through used cells in Aspose.Cells | process each visible sheet sequentially and then save the workbook in C# | check worksheet IsVisible property before reading cells with Aspose.Cells
// Tags: visible worksheet iteration Aspose.Cells C# | hidden sheet exclusion Aspose.Cells | used cells traversal Aspose.Cells | cell string trimming Aspose.Cells | workbook save after modification Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel file with Aspose.Cells, iterates through all worksheets, processes only those where IsVisible is true, trims whitespace from string cell values while traversing used rows and columns, writes the results to the console, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        // Paths to the source and destination Excel files
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook (using the standard load rule)
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Process only visible worksheets
                if (!sheet.IsVisible)
                    continue;

                // Iterate through all used cells in the current visible sheet
                Cells cells = sheet.Cells;
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];
                        if (cell == null || cell.Value == null)
                            continue;

                        // Retrieve the cell value as a string
                        string cellValue = cell.StringValue;

                        // Example processing: trim whitespace from string values
                        if (!string.IsNullOrEmpty(cellValue))
                        {
                            string processedValue = cellValue.Trim();

                            // (Optional) Perform additional logic with processedValue here
                            // For demonstration, we simply write it to the console
                            Console.WriteLine($"Sheet: {sheet.Name}, Row: {row + 1}, Column: {col + 1}, Value: {processedValue}");
                        }
                    }
                }
            }

            // Save the workbook after processing (using the standard save rule)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display an error message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
