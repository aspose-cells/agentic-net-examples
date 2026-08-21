// Title: Force Recalculation of Conditional Formatting Formulas with Aspose.Cells for .NET
// Description: Creates a workbook, fills cells A1:A5 with numbers, adds a conditional formatting rule that colors cells > 10 red, invokes Workbook.CalculateFormula to refresh any conditional formulas, and saves the file as an .xlsx document.
// Keywords: Aspose.Cells | Workbook.CalculateFormula | conditional formatting | C# .NET | recalculate formulas | Excel automation | force formula evaluation | Aspose.Cells conditional formatting | calculate after formatting
// Common Searches: Aspose.Cells calculate formula after conditional formatting | How to refresh conditional formatting in Aspose.Cells .NET | Workbook.CalculateFormula required for conditional formatting | Force conditional formatting update in C# Aspose.Cells | Reevaluate conditional formatting rules programmatically
// Developer Intent: Recalculate workbook formulas so that conditional formatting rules are applied correctly before the file is saved.
// Use Cases: Highlight cells with values greater than 10 in red and ensure the visual style reflects the latest data. | Generate a financial report where conditional formatting depends on calculated totals that must be refreshed prior to export. | Automate Excel creation in a .NET service, add data and formatting, then call CalculateFormula to guarantee accurate visual output.
// AI Prompts: Write C# code using Aspose.Cells to apply a conditional formatting rule for values > 20 and recalculate formulas before saving. | Explain why Workbook.CalculateFormula is necessary after adding conditional formatting in Aspose.Cells and provide a concise example. | Show how to combine multiple conditional formatting rules and trigger a single Workbook.CalculateFormula call to update all formulas.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, fills cells A1:A5 with numbers, adds a conditional formatting rule that colors cells > 10 red, invokes Workbook.CalculateFormula to refresh any conditional formulas, and saves the file as an .xlsx document.
    public class ConditionalFormattingCalculateDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data
                cells["A1"].PutValue(5);
                cells["A2"].PutValue(12);
                cells["A3"].PutValue(20);
                cells["A4"].PutValue(8);
                cells["A5"].PutValue(15);

                // Add a conditional formatting that highlights cells with values > 10
                int cfIndex = sheet.ConditionalFormattings.Add(); // create conditional formatting collection
                FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];

                // Define the range A1:A5
                CellArea area = new CellArea
                {
                    StartRow = 0,
                    StartColumn = 0,
                    EndRow = 4,
                    EndColumn = 0
                };
                fcc.AddArea(area);

                // Add the condition (type: CellValue, operator: GreaterThan, formula1: "10")
                int conditionIdx = fcc.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "10", null);
                FormatCondition condition = fcc[conditionIdx];

                // Set the style for the condition (red background)
                Style style = workbook.CreateStyle();
                style.ForegroundColor = Color.Red;
                style.Pattern = BackgroundType.Solid;
                condition.Style = style;

                // Recalculate formulas to ensure any conditional formulas are evaluated
                workbook.CalculateFormula(); // lifecycle rule: calculate

                // Save the workbook (lifecycle rule: save)
                workbook.Save("ConditionalFormattingCalculateDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ConditionalFormattingCalculateDemo.Run();
        }
    }
}
