// Title: Add conditional formatting and force formula recalculation in an Aspose.Cells workbook with C#
// AI Prompts: Create a new Workbook, fill cells A1‑A3 with values, add a conditional formatting rule that highlights values between 10 and 20 in red, call Workbook.CalculateFormula, and save the file. | Open an existing .xlsx file, modify its conditional formatting to use a CellValue Between condition, then programmatically trigger a full formula recalculation using Workbook.CalculateFormula before saving. | Write C# code that sets up a conditional formatting range, applies a style, invokes CalculateFormula to update dependent formulas, and writes the workbook to disk.
// Common Searches: Aspose.Cells C# how to recalculate formulas after adding conditional formatting | force Workbook.CalculateFormula after conditional formatting rule in Aspose.Cells | C# example for applying cell value between conditional formatting and updating formulas | recalculate dependent formulas when conditional formatting changes using Aspose.Cells | Aspose.Cells calculate formulas programmatically after conditional format changes
// Tags: conditional formatting with Aspose.Cells C# | Workbook.CalculateFormula after conditional formatting | apply cell value between condition Aspose.Cells | recalculate formulas in Excel workbook using Aspose.Cells | save workbook after conditional formatting C#

using System;
using Aspose.Cells;

namespace ConditionalFormattingRecalcDemo
{
    // // Demonstrates creating a workbook, populating cells, adding a conditional formatting rule that highlights values between 10 and 20, invoking Workbook.CalculateFormula to recalculate any dependent formulas, and saving the result as an .xlsx file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data
            cells["A1"].PutValue(5);
            cells["A2"].PutValue(15);
            cells["A3"].PutValue(25);

            // Add a conditional formatting rule: highlight values between 10 and 20
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];

            // Define the range to which the rule applies (A1:A3)
            CellArea area = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = 2,
                EndColumn = 0
            };
            fcc.AddArea(area);

            // Add the condition
            int conditionIdx = fcc.AddCondition(FormatConditionType.CellValue, OperatorType.Between, "10", "20");
            FormatCondition condition = fcc[conditionIdx];
            // Set a simple style (red background)
            condition.Style.BackgroundColor = System.Drawing.Color.Red;

            // After setting up conditional formatting, recalculate formulas
            // This ensures any formulas that depend on the conditional formatting are updated
            workbook.CalculateFormula();

            // Save the workbook (lifecycle: save)
            workbook.Save("ConditionalFormattingRecalc.xlsx");
        }
    }
}
