// Title: Trim whitespace from all string cells in a CSV file with Aspose.Cells for .NET and save the cleaned CSV
// Common Searches: C# Aspose.Cells remove leading and trailing spaces from CSV cells | How to clean whitespace in CSV data using Aspose.Cells workbook | Save trimmed CSV after processing cells with Aspose.Cells .NET | Iterate through cells in Aspose.Cells to trim string values in CSV | Aspose.Cells CSV whitespace cleanup example in C#
// Tags: trim whitespace cells Aspose.Cells CSV | CSV data cleaning with Aspose.Cells .NET | remove spaces from string cells C# | save cleaned workbook as CSV Aspose.Cells | process CSV workbook cell values Aspose.Cells

using System;
using Aspose.Cells;

namespace CsvCleaner
{
    // The example loads a CSV file into an Aspose.Cells Workbook, iterates over each used cell, trims leading and trailing whitespace from string cells, and then saves the cleaned data back to a CSV file.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source CSV file
            string inputPath = "input.csv";

            // Load the CSV file into a workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets (CSV typically has one sheet)
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the used range of cells
                Cells cells = sheet.Cells;
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                // Loop through each cell in the used range
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];

                        // Check if the cell contains a string value
                        if (cell.Type == CellValueType.IsString)
                        {
                            // Trim leading and trailing whitespace and write back
                            string trimmed = cell.StringValue.Trim();
                            cell.PutValue(trimmed);
                        }
                    }
                }
            }

            // Path to the cleaned CSV file
            string outputPath = "output.csv";

            // Save the cleaned workbook as CSV
            workbook.Save(outputPath, SaveFormat.CSV);
        }
    }
}
