using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsConditionalTextboxDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Put a numeric value in cell A1 (this will be the linked cell)
                worksheet.Cells["A1"].PutValue(75);

                // Add a textbox to the worksheet
                int textboxIndex = worksheet.TextBoxes.Add(2, 1, 160, 30);
                TextBox textbox = worksheet.TextBoxes[textboxIndex];

                // Set initial text and link the textbox to cell A1
                textbox.Text = "Current Value:";
                textbox.LinkedCell = "$A$1";

                // -----------------------------------------------------------------
                // Define conditional formatting on the linked cell (A1)
                // If the value > 50 -> light green background
                // If the value <= 50 -> light coral background
                // -----------------------------------------------------------------

                // Create a new conditional formatting collection
                int cfIndex = worksheet.ConditionalFormattings.Add();
                FormatConditionCollection conditions = worksheet.ConditionalFormattings[cfIndex];

                // Define the range that the formatting will apply to (cell A1)
                CellArea area = new CellArea
                {
                    StartRow = 0,   // Row 0 = A
                    EndRow = 0,
                    StartColumn = 0, // Column 0 = 1
                    EndColumn = 0
                };
                conditions.AddArea(area);

                // Condition 1: value greater than 50
                int condIdx1 = conditions.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "50", null);
                FormatCondition cond1 = conditions[condIdx1];
                cond1.Style.BackgroundColor = Color.LightGreen;

                // Condition 2: value less than or equal to 50
                int condIdx2 = conditions.AddCondition(FormatConditionType.CellValue, OperatorType.LessOrEqual, "50", null);
                FormatCondition cond2 = conditions[condIdx2];
                cond2.Style.BackgroundColor = Color.LightCoral;

                // Save the workbook
                string outputPath = "ConditionalTextbox.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}