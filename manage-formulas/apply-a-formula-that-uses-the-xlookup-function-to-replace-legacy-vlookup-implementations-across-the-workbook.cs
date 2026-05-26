using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ReplaceVlookupWithXlookup
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Search only in formulas for the text "VLOOKUP"
            FindOptions findOptions = new FindOptions
            {
                LookInType = LookInType.OnlyFormulas,
                LookAtType = LookAtType.Contains
            };

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                Cell startCell = null;

                // Find each occurrence of VLOOKUP and replace it
                while (true)
                {
                    Cell foundCell = cells.Find("VLOOKUP", startCell, findOptions);
                    if (foundCell == null)
                        break;

                    string oldFormula = foundCell.Formula;
                    string newFormula = oldFormula.Replace("VLOOKUP", "XLOOKUP");

                    // Update the formula; value will be recalculated later
                    foundCell.Formula = newFormula;

                    // Continue searching after the current cell
                    startCell = foundCell;
                }
            }

            // Recalculate all formulas
            workbook.CalculateFormula();

            // Save the modified workbook
            workbook.Save(outputPath);
        }
    }
}