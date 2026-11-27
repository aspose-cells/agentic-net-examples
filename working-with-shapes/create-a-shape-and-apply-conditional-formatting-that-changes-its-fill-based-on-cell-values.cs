// Title: Create an oval shape and change its fill color with conditional formatting based on a worksheet cell using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set shape fill color from a worksheet cell value | How to apply conditional formatting logic to a shape in Aspose.Cells | Change oval fill dynamically based on cell B2 in a .NET workbook | Example of syncing shape fill with cell value using Aspose.Cells | C# code to evaluate a cell and update shape fill color in Excel
// Tags: add oval shape Aspose.Cells | conditional formatting cell value Aspose.Cells | shape fill color based on worksheet cell | evaluate cell value set shape fill C# | save workbook with shape conditional fill Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeConditionalFormattingDemo
{
    // The example creates a new workbook, inserts an oval shape, defines conditional formatting rules for cell B2, evaluates the cell value, and sets the oval's fill to light green or light coral accordingly before saving the file as an XLSX document.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // 2. Populate a cell that will drive the conditional formatting
                //    In this example we use cell B2 (row index 1, column index 1)
                cells["B2"].PutValue(75); // Change this value to test different outcomes

                // 3. Add an oval shape to the worksheet
                //    Parameters: topRow, top (pixel offset), leftColumn, left (pixel offset), height, width
                Oval oval = worksheet.Shapes.AddOval(2, 0, 2, 0, 100, 200);
                oval.Name = "DynamicFillOval";

                // 4. Create a conditional formatting rule for the cell B2
                //    The rule will colour the cell background light green when the value > 50,
                //    otherwise the background will be light coral.
                int cfIndex = worksheet.ConditionalFormattings.Add();
                FormatConditionCollection cfCollection = worksheet.ConditionalFormattings[cfIndex];

                // Define the area the rule applies to (only B2 in this case)
                CellArea area = new CellArea
                {
                    StartRow = 1,
                    EndRow = 1,
                    StartColumn = 1,
                    EndColumn = 1
                };
                cfCollection.AddArea(area);

                // Add a CellValue condition: B2 > 50
                int conditionIndex = cfCollection.AddCondition(FormatConditionType.CellValue);
                FormatCondition condition = cfCollection[conditionIndex];
                condition.Operator = OperatorType.GreaterThan;
                condition.Formula1 = "50";

                // Set the style that will be applied when the condition is true
                condition.Style.BackgroundColor = Color.LightGreen;

                // Optional: set a style for the false case using a second rule
                int falseConditionIndex = cfCollection.AddCondition(FormatConditionType.CellValue);
                FormatCondition falseCondition = cfCollection[falseConditionIndex];
                falseCondition.Operator = OperatorType.LessOrEqual;
                falseCondition.Formula1 = "50";
                falseCondition.Style.BackgroundColor = Color.LightCoral;

                // 5. Evaluate the cell value and adjust the shape's fill accordingly
                //    This mimics "conditional formatting on the shape" by using the same logic.
                double cellValue = cells["B2"].DoubleValue;
                if (cellValue > 50)
                {
                    oval.Fill.FillType = FillType.Solid;
                    oval.Fill.SolidFill.Color = Color.LightGreen;
                }
                else
                {
                    oval.Fill.FillType = FillType.Solid;
                    oval.Fill.SolidFill.Color = Color.LightCoral;
                }

                // 6. Save the workbook
                string outputPath = "ShapeConditionalFormattingDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                Console.WriteLine($"Cell B2 value = {cellValue}. Shape fill set to {(cellValue > 50 ? "LightGreen" : "LightCoral")}.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
