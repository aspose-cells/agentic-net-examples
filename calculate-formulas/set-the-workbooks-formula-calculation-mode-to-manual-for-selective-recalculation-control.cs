// Title: Set manual formula calculation mode for an Aspose.Cells workbook in C#
// AI Prompts: Write C# code that creates an Aspose.Cells workbook, assigns CalcModeType.Manual to FormulaSettings.CalculationMode, adds sample cells and a formula, then saves the file. | Show how to enable selective recalculation by configuring manual calculation mode in an Aspose.Cells workbook using the .NET API.
// Common Searches: how to enable manual calculation mode in Aspose.Cells with C# | Aspose.Cells workbook prevent automatic formula evaluation .NET | set CalcModeType.Manual for selective recalculation in Aspose.Cells | C# example of disabling auto formula recalculation in Excel file using Aspose.Cells | manual formula calculation setting Aspose.Cells API
// Tags: Aspose.Cells manual formula calculation | C# set workbook calculation mode | prevent auto formula evaluation Aspose.Cells | Aspose.Cells calculation mode setting | partial formula evaluation Aspose.Cells

using Aspose.Cells;

// Creates a new workbook, switches its formula calculation mode to Manual for selective recalculation, adds sample data and a simple formula, and saves the workbook as ManualCalcMode.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set the calculation mode to Manual for selective recalculation
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

        // Example data and a formula (optional)
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["A2"].PutValue(10);
        sheet.Cells["A3"].Formula = "=A1+A2";

        // Save the workbook
        workbook.Save("ManualCalcMode.xlsx");
    }
}
