// Title: Add a COUNT subtotal for column K in an Excel workbook with Aspose.Cells for .NET without using outline levels
// AI Prompts: Write C# code that uses Aspose.Cells to locate the last populated row in the first worksheet, then inserts a label in column J and a =COUNT(K2:K{lastRow}) formula in column K two rows below the data, ensuring no outline hierarchy is created. | Generate a .NET snippet that loads an existing .xlsx file, adds a numeric count subtotal for column K by placing the formula after the data range, and saves the workbook, all with Aspose.Cells and without creating outline levels.
// Common Searches: aspocells count values in column K without outline | c# add subtotal row at end of data using Aspose.Cells | how to insert COUNT formula programmatically in Excel with Aspose.Cells .NET | Aspose.Cells place formula two rows below last row | calculate numeric count of a column in .xlsx using Aspose.Cells C#
// Tags: Aspose.Cells insert COUNT formula column K | C# add subtotal row without outline levels | Excel numeric count subtotal using Aspose.Cells | determine last data row Aspose.Cells | save workbook after adding formula Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an input.xlsx file, checks its existence, and opens the first worksheet. It determines the last row containing data, then writes a descriptive label in column J and a =COUNT(K2:K{lastRow}) formula in column K two rows below the data range, without creating any outline levels. The workbook is saved as output.xlsx, and the code includes basic error handling and directory creation.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Access the first worksheet
            var sheet = workbook.Worksheets[0];

            // Determine the last row that contains data
            int lastRow = sheet.Cells.MaxDataRow;

            // If there is data, add a formula that counts numeric values in column K (index 10)
            if (lastRow >= 1) // assuming row 0 contains headers
            {
                // Place the result two rows below the last data row
                int resultRow = lastRow + 2;
                // Excel rows are 1‑based, so add 1 to indices
                string startCell = $"K{2}";
                string endCell = $"K{lastRow + 1}";
                sheet.Cells[resultRow, 10].Formula = $"=COUNT({startCell}:{endCell})";
                sheet.Cells[resultRow, 9].PutValue("Count of numeric values in K:");
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
