// Title: Load an Excel workbook and verify automatic row heights for each worksheet using Aspose.Cells in C#
// Common Searches: how to read row height values for all rows in a worksheet with Aspose.Cells C# | check if a row uses automatic height after loading a workbook using Aspose.Cells | iterate through every worksheet in a workbook and output row height information with Aspose.Cells | determine rows with default (zero) height in an Excel file using Aspose.Cells C# | verify automatic row height calculation across multiple sheets with Aspose.Cells
// Tags: load workbook Aspose.Cells C# | iterate worksheets Aspose.Cells | read row height property Aspose.Cells | detect auto row height Aspose.Cells | check rows with zero height Aspose.Cells | validate automatic row height Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads 'input.xlsx' using Aspose.Cells, iterates through each worksheet, finds the last used row, and examines the Height property of every row. A height of zero is interpreted as automatically calculated. The program prints the sheet name, row index, whether the height is auto, and the actual height, while handling missing files and unexpected exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook with default load options
            var loadOptions = new LoadOptions();
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Determine the last used row in the sheet
                int maxRow = sheet.Cells.MaxDataRow;

                // Loop through all rows up to the last used row
                for (int rowIndex = 0; rowIndex <= maxRow; rowIndex++)
                {
                    Row row = sheet.Cells.Rows[rowIndex];

                    // Since Row.IsCustomHeight is not available in this version,
                    // we treat the height as automatically calculated when Height is zero.
                    bool isAutoHeight = row.Height == 0;

                    // Output verification result
                    Console.WriteLine($"Sheet: {sheet.Name}, Row: {rowIndex + 1}, AutoHeight: {isAutoHeight}, Height: {row.Height}");
                }
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
