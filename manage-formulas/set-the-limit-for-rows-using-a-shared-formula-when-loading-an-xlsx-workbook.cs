using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the XLSX workbook (you can customize LoadOptions if needed)
        LoadOptions loadOptions = new LoadOptions();
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Set the maximum number of rows that a shared formula can span
        // Adjust this value according to your requirements
        workbook.Settings.MaxRowsOfSharedFormula = 500;

        // Recalculate formulas so that any shared formulas respect the new limit
        workbook.CalculateFormula();

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}