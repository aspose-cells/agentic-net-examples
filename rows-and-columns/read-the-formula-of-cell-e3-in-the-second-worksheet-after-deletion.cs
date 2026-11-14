// Title: Read the formula from cell E3 on the second worksheet after deleting the first row with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# get formula of a cell after deleting rows | how to read cell E3 formula on second sheet after row removal using Aspose.Cells | C# Aspose.Cells retrieve formula from specific cell after worksheet row deletion | read Excel formula from cell after deleting first row with Aspose.Cells .NET | Aspose.Cells example delete first row then access cell formula
// Tags: Aspose.Cells read cell formula .NET | delete rows then access cell formula Aspose.Cells | second worksheet cell E3 formula Aspose.Cells | C# Excel row deletion impact on formulas | Aspose.Cells workbook manipulation read formula

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook, selects the second worksheet, deletes the first row, then reads and prints the formula stored in cell E3, with error handling for missing files and other exceptions.
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
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Access the second worksheet (index 1)
            Worksheet worksheet = workbook.Worksheets[1];

            // Delete the first row in this worksheet (startRow = 0, totalRows = 1)
            worksheet.Cells.DeleteRows(0, 1);

            // Retrieve the cell at E3 (row 2, column 4)
            Cell targetCell = worksheet.Cells["E3"];

            // Read the formula stored in the cell
            string formula = targetCell.Formula;

            // Output the formula to the console
            Console.WriteLine("Formula in E3 of the second worksheet after deletion: " + formula);
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display an error message
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
