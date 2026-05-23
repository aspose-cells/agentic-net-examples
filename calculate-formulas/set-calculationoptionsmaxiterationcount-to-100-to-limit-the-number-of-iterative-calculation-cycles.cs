using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Enable iterative calculation and limit iterations to 100
        workbook.Settings.FormulaSettings.EnableIterativeCalculation = true;
        workbook.Settings.FormulaSettings.MaxIteration = 100;

        // Save the workbook
        workbook.Save("IterativeCalculation.xlsx");
    }
}