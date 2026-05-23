using System;
using Aspose.Cells;

namespace ReplaceFormulaReferences
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Replace all occurrences of the named range "OldRange" with "NewRange"
            // This works for formulas, cell values, and any text containing the placeholder.
            int replacedCount = workbook.Replace("OldRange", "NewRange");

            // Optionally, you can output how many replacements were made
            Console.WriteLine($"Replaced {replacedCount} occurrences of 'OldRange' with 'NewRange'.");

            // Save the modified workbook to a new file
            workbook.Save("output.xlsx");
        }
    }
}