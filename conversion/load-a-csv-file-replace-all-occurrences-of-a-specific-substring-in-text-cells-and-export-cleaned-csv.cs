// Title: Replace a specific substring in every text cell of a CSV and export the cleaned file using Aspose.Cells for .NET (C#)
// AI Prompts: Load a CSV into an Aspose.Cells Workbook, scan all string cells, replace each occurrence of a given substring with a new value, and write the result to a new CSV file in C#. | Using Aspose.Cells for .NET, iterate over the used range of a worksheet, perform bulk string substitution on text cells, and save the modified workbook as CSV. | Write C# code that opens a CSV with Aspose.Cells, substitutes "OldValue" with "NewValue" in every text cell, and exports the cleaned data back to CSV.
// Common Searches: Aspose.Cells C# replace substring in CSV cells and save | How to perform bulk text replacement in a CSV using Aspose.Cells .NET | C# iterate through all cells of a CSV workbook and modify string values with Aspose.Cells | Replace specific text in CSV file using Aspose.Cells for .NET | Save modified CSV after string substitution with Aspose.Cells
// Tags: csv bulk string replace Aspose.Cells | iterate used range Aspose.Cells C# | export cleaned CSV Aspose.Cells | text cell processing Aspose.Cells .NET | substring substitution in worksheet cells

using System;
using Aspose.Cells;

namespace CsvCleaner
{
    // // Loads a CSV with Aspose.Cells, iterates over all used cells, replaces a target substring in string cells, and saves the cleaned data to a new CSV file.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source CSV file
            string inputPath = "input.csv";

            // Path for the cleaned CSV output
            string outputPath = "output.csv";

            // Substring to find and its replacement
            string oldSubstring = "OldValue";
            string newSubstring = "NewValue";

            // Load the CSV file into a Workbook object
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

                        // Process only text (string) cells
                        if (cell.Type == CellValueType.IsString)
                        {
                            string original = cell.StringValue;
                            // Replace the target substring
                            string replaced = original.Replace(oldSubstring, newSubstring);
                            // Update the cell only if a change occurred
                            if (!original.Equals(replaced))
                            {
                                cell.PutValue(replaced);
                            }
                        }
                    }
                }
            }

            // Save the modified workbook back to CSV format
            workbook.Save(outputPath, SaveFormat.CSV);
        }
    }
}
