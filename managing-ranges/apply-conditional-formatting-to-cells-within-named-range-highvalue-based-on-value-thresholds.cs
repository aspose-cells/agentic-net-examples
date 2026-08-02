// Title: Conditional formatting for named range 'HighValue' (A1:A10) in Aspose.Cells C#
// Description: Creates a workbook, fills A1:A10 with values 0‑180, defines the named range HighValue, and applies two conditional formats – values > 100 get a red fill, values < 50 get a light‑blue fill – then saves the file.
// Keywords: Aspose.Cells | C# | conditional formatting | named range | HighValue | Excel automation | value thresholds | greater than 100 red | less than 50 light blue | CellArea | FormatCondition | Excel report styling
// Common Searches: Aspose.Cells C# conditional formatting named range | how to highlight cells >100 red using Aspose.Cells | apply light blue background to cells <50 Aspose.Cells | C# create named range HighValue in Aspose.Cells | conditional formatting multiple rules Aspose.Cells .NET
// Developer Intent: Add two value‑based conditional‑formatting rules to the named range HighValue so cells >100 turn red and cells <50 turn light‑blue.
// Use Cases: Flag sales figures that surpass a target by coloring them red. | Show inventory items below reorder level with a light‑blue background. | Separate high‑risk and low‑risk KPI values in a single column for quick visual analysis. | Create a dashboard where thresholds are highlighted automatically in Excel reports.
// AI Prompts: Write C# code using Aspose.Cells to create a named range "HighValue" for A1:A10 and apply red fill for values >100 and light‑blue fill for values <50. | Refactor the sample into a method that accepts a worksheet, a named range, and customizable threshold values for conditional formatting. | Explain how to add conditional formatting without overwriting existing cell styles in an Aspose.Cells workbook. | Generate unit tests that verify the conditional formatting rules are applied correctly to the HighValue range.

using System;
using System.Drawing;
using Aspose.Cells;

namespace ConditionalFormattingDemo
{
    // Creates a workbook, fills A1:A10 with values 0‑180, defines the named range HighValue, and applies two conditional formats – values > 100 get a red fill, values < 50 get a light‑blue fill – then saves the file.
    class ConditionalFormattingExample
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data in column A (rows 1 to 10)
                for (int i = 0; i < 10; i++)
                {
                    sheet.Cells[i, 0].PutValue(i * 20); // values: 0,20,40,...,180
                }

                // Build the CellArea that represents the range A1:A10
                CellArea rangeArea = new CellArea
                {
                    StartRow = 0,
                    EndRow = 9,
                    StartColumn = 0,
                    EndColumn = 0
                };

                // Add a new conditional formatting collection to the worksheet
                int cfIndex = sheet.ConditionalFormattings.Add();
                FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];

                // Associate the range with the conditional formatting
                fcc.AddArea(rangeArea);

                // Condition 1: Highlight cells with values greater than 100 (red background)
                int condIdx1 = fcc.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "100", null);
                FormatCondition fc1 = fcc[condIdx1];
                fc1.Style.BackgroundColor = Color.Red;

                // Condition 2: Highlight cells with values less than 50 (light blue background)
                int condIdx2 = fcc.AddCondition(FormatConditionType.CellValue, OperatorType.LessThan, "50", null);
                FormatCondition fc2 = fcc[condIdx2];
                fc2.Style.BackgroundColor = Color.LightBlue;

                // Determine output file path
                string outputFile = "ConditionalFormattingHighValue.xlsx";

                // Save the workbook
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
