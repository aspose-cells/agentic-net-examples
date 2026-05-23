using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set calculation mode to SemiAutomatic (AutomaticExceptTable)
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.AutomaticExceptTable;

        // Add sample data and a formula
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["A2"].PutValue(20);
        sheet.Cells["A3"].Formula = "=A1+A2";

        // Save the workbook
        workbook.Save("SemiAutomaticMode.xlsx");
    }
}