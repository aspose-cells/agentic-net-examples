using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsConditionalFormattingDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Sample data: mix of positive and negative numbers
                sheet.Cells["A1"].PutValue(25);
                sheet.Cells["A2"].PutValue(-10);
                sheet.Cells["A3"].PutValue(40);
                sheet.Cells["A4"].PutValue(-5);
                sheet.Cells["A5"].PutValue(15);

                // Add a conditional formatting collection to the worksheet
                int cfIndex = sheet.ConditionalFormattings.Add();
                FormatConditionCollection cfCollection = sheet.ConditionalFormattings[cfIndex];

                // Define the range to which the formatting will be applied (A1:A5)
                CellArea area = new CellArea
                {
                    StartRow = 0,
                    EndRow = 4,
                    StartColumn = 0,
                    EndColumn = 0
                };
                cfCollection.AddArea(area);

                // Add a condition: cell value less than 0 (negative numbers)
                int conditionIndex = cfCollection.AddCondition(
                    FormatConditionType.CellValue,
                    OperatorType.LessThan,
                    "0",
                    null);

                FormatCondition condition = cfCollection[conditionIndex];

                // Set the font color for negative numbers (using a solid red color)
                condition.Style.Font.Color = Color.Red;

                // Save the workbook
                string outputPath = "NegativeFontColor_Red.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}