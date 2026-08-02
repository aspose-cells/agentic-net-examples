using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set calculation mode to AutomaticExceptTable.
        // This mode behaves like a “semi‑automatic” mode where only dependent cells are recalculated.
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.AutomaticExceptTable;

        // Sample data and formula to demonstrate the setting
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["A2"].PutValue(10);
        sheet.Cells["B1"].Formula = "=A1+A2";

        // Save the workbook
        workbook.Save("SemiAutomaticMode.xlsx");
    }
}