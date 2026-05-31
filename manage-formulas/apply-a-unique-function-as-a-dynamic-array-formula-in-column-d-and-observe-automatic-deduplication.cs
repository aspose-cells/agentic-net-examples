using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data in column C (C2:C10) with some duplicate values
            string[] sampleData = { "Apple", "Orange", "Apple", "Banana", "Orange", "Grape", "Banana", "Kiwi", "Apple" };
            for (int i = 0; i < sampleData.Length; i++)
            {
                // Row index i+1 (starts at 2), column index 2 (C)
                cells[i + 1, 2].PutValue(sampleData[i]);
            }

            // Apply the UNIQUE function as a dynamic array formula in D2
            // The formula will spill into as many rows as needed, automatically removing duplicates
            Cell targetCell = cells[1, 3]; // D2
            string uniqueFormula = "=UNIQUE(C2:C10)";
            targetCell.SetDynamicArrayFormula(uniqueFormula, new FormulaParseOptions(), calculateValue: true);

            // Refresh dynamic array formulas to ensure the spill range is up‑to‑date
            workbook.RefreshDynamicArrayFormulas(calculate: true);

            // Optionally calculate the workbook (not strictly needed because RefreshDynamicArrayFormulas with calculate:true already does it)
            workbook.CalculateFormula();

            // Output the results of the spilled range (D2 and below)
            Console.WriteLine("Unique values from column C (spilled into column D):");
            int row = 1; // start at D2 (row index 1)
            while (true)
            {
                Cell c = cells[row, 3]; // column D (index 3)
                // Stop when the cell is null or contains no value
                if (c == null || string.IsNullOrEmpty(c.StringValue))
                    break;

                Console.WriteLine(c.StringValue);
                row++;
            }

            // Save the workbook (lifecycle: save)
            string outputPath = "UniqueDynamicArrayDemo.xlsx";
            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}