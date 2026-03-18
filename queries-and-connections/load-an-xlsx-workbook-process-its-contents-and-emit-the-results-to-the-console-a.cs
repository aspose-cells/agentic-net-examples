using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source Excel file
        string inputPath = "input.xlsx";

        // Path to the output text file
        string outputTxt = "output.txt";

        // Load the workbook using the string constructor (load rule)
        Workbook workbook = new Workbook(inputPath);

        // Open a StreamWriter to write results to the text file
        using (StreamWriter writer = new StreamWriter(outputTxt))
        {
            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Write worksheet name to console and text file
                Console.WriteLine($"Worksheet: {sheet.Name}");
                writer.WriteLine($"Worksheet: {sheet.Name}");

                // Determine the used range of the worksheet
                int maxRow = sheet.Cells.MaxDataRow;
                int maxCol = sheet.Cells.MaxDataColumn;

                // Iterate over all used cells
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = sheet.Cells[row, col];
                        if (cell.Value != null)
                        {
                            string address = cell.Name;          // e.g., "A1"
                            string value = cell.Value.ToString(); // cell content

                            // Output to console and text file
                            Console.WriteLine($"{address}: {value}");
                            writer.WriteLine($"{address}: {value}");
                        }
                    }
                }

                // Add a blank line between worksheets
                Console.WriteLine();
                writer.WriteLine();
            }
        }

        // (Optional) Save the workbook if any modifications were made
        // workbook.Save("processed.xlsx");
    }
}