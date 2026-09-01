// Title: How to convert an Excel workbook to CSV and keep formulas as plain text with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file using Aspose.Cells, substitutes every formula cell with its formula string, and saves the result as a CSV file. | Demonstrate iterating over the used range of a worksheet, converting formulas to text, and exporting the workbook to CSV in a .NET application.
// Common Searches: Aspose.Cells C# export Excel to CSV preserving formula text | keep Excel formulas when converting to CSV using Aspose.Cells | replace formulas with their string representation before saving as CSV in .NET
// Tags: Aspose.Cells CSV export with formulas as text | C# replace Excel formula with its string value | iterate over used range Aspose.Cells | save workbook as CSV using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example loads an input.xlsx workbook with Aspose.Cells, walks through all used cells, replaces any cell containing a formula with the formula's text, and then saves the workbook as output.csv, ensuring that formulas are retained as plain text in the CSV output.
    public class WorkbookToCsvPreserveFormulas
    {
        public static void Main(string[] args)
        {
            try
            {
                // Path to the source Excel workbook
                string sourcePath = "input.xlsx";

                // Path for the resulting CSV file
                string csvPath = "output.csv";

                // Verify that the source file exists
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Error: The source file \"{sourcePath}\" was not found.");
                    return;
                }

                // Load the workbook from the file
                Workbook workbook = new Workbook(sourcePath);

                // Access the first worksheet (adjust if needed)
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Determine the used range of the worksheet
                int maxRow = cells.MaxDataRow;
                int maxColumn = cells.MaxDataColumn;

                // Iterate through all used cells
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxColumn; col++)
                    {
                        // Get the current cell
                        Cell cell = cells[row, col];

                        // If the cell contains a formula, replace its value with the formula text
                        if (!string.IsNullOrEmpty(cell.Formula))
                        {
                            // Preserve the formula as a plain string
                            cell.PutValue(cell.Formula);
                        }
                    }
                }

                // Save the modified workbook as CSV; formulas are now stored as text strings
                workbook.Save(csvPath, SaveFormat.Csv);

                Console.WriteLine($"Workbook converted to CSV with formulas preserved as text: {csvPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
