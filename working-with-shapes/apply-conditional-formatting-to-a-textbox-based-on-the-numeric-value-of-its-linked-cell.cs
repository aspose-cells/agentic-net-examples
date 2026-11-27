// Title: How to apply conditional formatting to a linked TextBox shape in an Excel file using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set TextBox fill color based on linked cell value | Conditional formatting of a TextBox shape linked to a worksheet cell using Aspose.Cells | Change Aspose.Cells TextBox appearance when cell value exceeds a threshold | Synchronize TextBox styling with cell A1 > 50 in Aspose.Cells .NET | Apply style to linked TextBox when worksheet cell meets condition in C#
// Tags: conditional formatting linked textbox Aspose.Cells | text box shape styling based on cell value C# | apply fill color to Aspose.Cells TextBox | synchronize shape appearance with cell condition .NET | dynamic textbox formatting Excel Aspose.Cells | evaluate cell value and update TextBox style C#

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ConditionalFormattingTextBoxDemo
{
    // The example creates a workbook, writes a numeric value to cell A1, adds a TextBox linked to that cell, defines a conditional formatting rule (value > 50) that sets a red background and white bold font, checks the cell value at runtime, and mirrors the formatting onto the TextBox by updating its fill and font properties before saving the file as XLSX.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to output file
            const string outputPath = "ConditionalFormattingTextBoxDemo.xlsx";

            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // 2. Populate a linked cell with a numeric value
                //    This cell will drive the conditional formatting of the TextBox
                const string linkedCellName = "A1";
                Cell linkedCell = cells[linkedCellName];
                linkedCell.PutValue(75); // Change this value to test different outcomes

                // 3. Add a TextBox shape to the worksheet
                //    Parameters: upper left row, upper left column, upper left offset (pixels),
                //                upper left offset (pixels), width (pixels), height (pixels)
                int textboxIndex = worksheet.TextBoxes.Add(2, 1, 200, 100);
                TextBox textBox = worksheet.TextBoxes[textboxIndex];
                textBox.Text = "Linked Value: " + linkedCell.Value;
                textBox.Font.Size = 12;
                textBox.Font.Color = Color.Black;
                textBox.Placement = PlacementType.MoveAndSize;

                // Link the TextBox to the cell so that its text updates automatically
                textBox.LinkedCell = linkedCellName;

                // 4. Define conditional formatting on the linked cell
                //    When the cell value is greater than 50, apply a red background and white font.
                int cfIndex = worksheet.ConditionalFormattings.Add();
                FormatConditionCollection fcs = worksheet.ConditionalFormattings[cfIndex];

                // Define the area the conditional formatting applies to (only the linked cell)
                CellArea area = new CellArea
                {
                    StartRow = linkedCell.Row,
                    EndRow = linkedCell.Row,
                    StartColumn = linkedCell.Column,
                    EndColumn = linkedCell.Column
                };
                fcs.AddArea(area);

                // Add a condition: CellValue > 50
                int conditionIndex = fcs.AddCondition(FormatConditionType.CellValue);
                FormatCondition condition = fcs[conditionIndex];
                condition.Operator = OperatorType.GreaterThan;
                condition.Formula1 = "50";

                // Set the style that will be applied when the condition is true
                condition.Style.BackgroundColor = Color.Red;
                condition.Style.Font.Color = Color.White;
                condition.Style.Font.IsBold = true;

                // 5. Recalculate formulas (not strictly needed here but good practice)
                workbook.CalculateFormula();

                // 6. Synchronize TextBox appearance with the conditional formatting result
                //    If the linked cell meets the condition, apply the same style to the TextBox.
                bool conditionMet = linkedCell.GetStyle().ForegroundColor.Equals(Color.White) ||
                                    linkedCell.GetStyle().BackgroundColor.Equals(Color.Red);

                // Alternative robust check: evaluate the condition manually
                if (linkedCell.DoubleValue > 50)
                {
                    // Apply background fill
                    FillFormat fill = textBox.Fill;
                    fill.SolidFill.Color = Color.Red;

                    // Apply font formatting
                    textBox.Font.Color = Color.White;
                    textBox.Font.IsBold = true;
                }
                else
                {
                    // Default appearance when condition is not met
                    FillFormat fill = textBox.Fill;
                    fill.SolidFill.Color = Color.LightGray;

                    textBox.Font.Color = Color.Black;
                    textBox.Font.IsBold = false;
                }

                // 7. Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
