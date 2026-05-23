using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the source workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Calculate all formulas in the workbook so that their results are up‑to‑date
        workbook.CalculateFormula();

        // Replace formulas with their calculated values for every worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.Cells.RemoveFormulas();
        }

        // Save the workbook as CSV – formulas are now static values
        workbook.Save("output.csv", SaveFormat.Csv);

        Console.WriteLine("Workbook has been exported to CSV with formulas evaluated.");
    }
}