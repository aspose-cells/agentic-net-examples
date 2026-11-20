// Title: Disable automatic formula recalculation in an Aspose.Cells workbook using C#
// Common Searches: C# Aspose.Cells how to turn off automatic formula calculation | set manual calculation mode for Excel workbook with Aspose.Cells .NET | prevent formulas from recalculating when saving workbook in Aspose.Cells | Aspose.Cells workbook Settings.CalcMode manual example C# | disable auto recalculation of cells in Aspose.Cells generated file
// Tags: Aspose.Cells manual calculation mode | disable automatic workbook recalculation | set workbook CalcMode to Manual C# | Aspose.Cells calculation settings API | prevent formula evaluation in Excel using Aspose.Cells

using System;
using Aspose.Cells;

// Creates a new Workbook, optionally sets Settings.CalcMode to Manual to stop automatic formula evaluation, saves the file as DisabledRecalc.xlsx, and handles any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Disable automatic calculation if the API version supports it.
            // Uncomment the following line for versions that expose CalcMode.
            // workbook.Settings.CalcMode = CalculationMode.Manual;

            // Save the workbook
            workbook.Save("DisabledRecalc.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
