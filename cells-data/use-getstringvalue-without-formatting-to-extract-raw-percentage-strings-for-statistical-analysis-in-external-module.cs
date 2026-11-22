// Title: How to extract the raw numeric string of a percentage cell from an Excel workbook using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells read percentage cell as raw numeric string without formatting | retrieve underlying value of a percentage formatted cell using Aspose.Cells | extract unformatted percentage value from Excel file with Aspose.Cells .NET | get raw cell value for statistical analysis from .xlsx using Aspose.Cells | Aspose.Cells cell.Value to string for percentage data extraction in C#
// Tags: Cell.Value raw numeric string Aspose.Cells | percentage cell extraction .NET | unformatted Excel data retrieval C# | statistical analysis raw percentage Aspose.Cells | read underlying value from formatted cell Aspose

using Aspose.Cells;
using System;
using System.IO;

// The example checks that 'input.xlsx' exists, loads it with Aspose.Cells, accesses the first worksheet, reads cell B2, and converts the cell's underlying numeric value (e.g., 0.25 for 25%) to a string using the Cell.Value property. The raw percentage string is printed for further statistical processing, with error handling for missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputFile = "input.xlsx";

        // Ensure the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Error: The file '{inputFile}' was not found.");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputFile);

            // Access the first worksheet (index 0)
            Worksheet worksheet = workbook.Worksheets[0];

            // Specify the cell that contains a percentage value (e.g., B2)
            Cell cell = worksheet.Cells["B2"];

            // Retrieve the raw underlying value of the cell (e.g., 0.25 for 25%)
            string rawPercentage = cell.Value?.ToString() ?? string.Empty;

            // Output the raw percentage string for further analysis
            Console.WriteLine($"Raw percentage value in B2: {rawPercentage}");

            // (Optional) Save the workbook if modifications were made
            // workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            // Handle any runtime exceptions gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
