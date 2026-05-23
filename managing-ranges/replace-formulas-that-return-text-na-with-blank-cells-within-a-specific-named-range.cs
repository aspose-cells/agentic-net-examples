using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsReplaceNA
{
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string rangeName = "MyRange";

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Retrieve the named range definition
                Name namedRange = workbook.Worksheets.Names[rangeName];
                if (namedRange == null)
                {
                    Console.WriteLine($"Named range \"{rangeName}\" not found.");
                    return;
                }

                // Parse the RefersTo string (e.g. =Sheet1!$A$1:$B$10)
                string refersTo = namedRange.RefersTo;
                if (refersTo.StartsWith("="))
                    refersTo = refersTo.Substring(1);

                int exclPos = refersTo.IndexOf('!');
                if (exclPos < 0)
                {
                    Console.WriteLine("Invalid range reference format.");
                    return;
                }

                string sheetName = refersTo.Substring(0, exclPos).Trim('\'');
                string address = refersTo.Substring(exclPos + 1);

                // Access the worksheet that contains the range
                Worksheet rangeSheet = workbook.Worksheets[sheetName];
                if (rangeSheet == null)
                {
                    Console.WriteLine($"Worksheet \"{sheetName}\" not found.");
                    return;
                }

                // Create a Range object from the address
                AsposeRange range = rangeSheet.Cells.CreateRange(address);

                // Ensure all formulas are calculated before inspection
                workbook.CalculateFormula();

                // Determine the bounds of the range
                int startRow = range.FirstRow;
                int startCol = range.FirstColumn;
                int endRow = startRow + range.RowCount - 1;
                int endCol = startCol + range.ColumnCount - 1;

                // Iterate through each cell in the range
                for (int row = startRow; row <= endRow; row++)
                {
                    for (int col = startCol; col <= endCol; col++)
                    {
                        Cell cell = rangeSheet.Cells[row, col];

                        // Replace cells whose evaluated value is "N/A"
                        if (cell.IsFormula && cell.StringValue == "N/A")
                        {
                            cell.PutValue(string.Empty);
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}