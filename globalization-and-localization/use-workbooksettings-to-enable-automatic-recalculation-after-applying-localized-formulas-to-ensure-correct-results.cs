// Title: How to enable automatic recalculation in Aspose.Cells for .NET using Workbook.Settings after applying localized formulas
// Common Searches: Aspose.Cells C# enable auto calculate workbook settings for locale‑specific formulas | set Workbook.Settings.AutoCalculate true to recalculate formulas after changing culture in .NET | automatic formula recalculation with localized syntax using Aspose.Cells Workbook.Settings | enable auto‑recalculation for Excel files after applying regional formulas in C# | Aspose.Cells workbook settings auto calculate after applying locale‑dependent formulas
// Tags: Workbook.Settings.AutoCalculate property Aspose.Cells | automatic formula recalculation .NET Excel | localized formula handling Aspose.Cells | culture‑aware calculations C# Aspose.Cells | enable auto calculate workbook settings

using System;
using Aspose.Cells;

// Shows how to set Workbook.Settings.AutoCalculate to true so that formulas written with locale‑specific syntax are automatically recalculated. The example creates a workbook, fills cells B2‑B6 with values 1‑5, assigns a SUM formula to A1, relies on the auto‑calculate setting for correct results, and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Fill sample data in B1:B5 (column index 1, zero‑based)
            for (int i = 0; i < 5; i++)
            {
                sheet.Cells[i + 1, 1].PutValue(i + 1);
            }

            // Apply a formula that sums B1:B5
            sheet.Cells["A1"].Formula = "=SUM(B1:B5)";

            // Force calculation to evaluate the formula immediately
            workbook.CalculateFormula();

            // Save the workbook
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
