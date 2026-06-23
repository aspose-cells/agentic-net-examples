using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsThemeValidation
{
    public class ValidateConditionalFormattingAfterTheme
    {
        public static void Run()
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // 2. Populate sample data in column A (A1:A5)
                for (int i = 0; i < 5; i++)
                {
                    sheet.Cells[i, 0].PutValue(10 + i * 10); // 10,20,30,40,50
                }

                // 3. Add a conditional formatting rule:
                //    Highlight cells with values between 15 and 35 (inclusive) with a red background.
                int cfIndex = sheet.ConditionalFormattings.Add();
                FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];

                // Define the range A1:A5
                CellArea area = new CellArea { StartRow = 0, EndRow = 4, StartColumn = 0, EndColumn = 0 };
                fcc.AddArea(area);

                // Add the condition
                int conditionIdx = fcc.AddCondition(FormatConditionType.CellValue, OperatorType.Between, "15", "35");
                FormatCondition fc = fcc[conditionIdx];
                Style cfStyle = workbook.CreateStyle();
                cfStyle.BackgroundColor = Color.Red;
                fc.Style = cfStyle;

                // 4. Capture the conditional formatting result for cell A3 (value 30) BEFORE applying the theme
                Cell targetCell = sheet.Cells["A3"]; // Row index 2
                ConditionalFormattingResult beforeResult = targetCell.GetConditionalFormattingResult();
                Style beforeStyle = beforeResult?.ConditionalStyle;

                Console.WriteLine("Before theme - Conditional style applied: " + (beforeStyle != null));
                if (beforeStyle != null)
                    Console.WriteLine("Before theme - Background color: " + beforeStyle.BackgroundColor);

                // 5. Define a custom theme (12 colors as required)
                Color[] customColors = new Color[]
                {
                    Color.White,          // Background1
                    Color.Black,          // Text1
                    Color.LightGray,      // Background2
                    Color.DarkGray,       // Text2
                    Color.Orange,         // Accent1
                    Color.Purple,         // Accent2
                    Color.Teal,           // Accent3
                    Color.Lime,           // Accent4
                    Color.Maroon,         // Accent5
                    Color.Navy,           // Accent6
                    Color.Blue,           // Hyperlink
                    Color.Red             // Followed Hyperlink
                };

                // 6. Apply the custom theme to the workbook
                workbook.CustomTheme("CustomDemoTheme", customColors);

                // 7. Capture the conditional formatting result for the same cell AFTER applying the theme
                ConditionalFormattingResult afterResult = targetCell.GetConditionalFormattingResult();
                Style afterStyle = afterResult?.ConditionalStyle;

                Console.WriteLine("After theme - Conditional style applied: " + (afterStyle != null));
                if (afterStyle != null)
                    Console.WriteLine("After theme - Background color: " + afterStyle.BackgroundColor);

                // 8. Validate that the conditional formatting still works (style should be unchanged)
                bool isStyleUnchanged = beforeStyle != null && afterStyle != null &&
                                        beforeStyle.BackgroundColor.ToArgb() == afterStyle.BackgroundColor.ToArgb();

                Console.WriteLine("Conditional formatting preserved after theme change: " + isStyleUnchanged);

                // 9. Save the workbook to verify the result manually if needed
                string outputPath = "ValidateConditionalFormattingAfterTheme.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine("Workbook saved to: " + Path.GetFullPath(outputPath));
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            ValidateConditionalFormattingAfterTheme.Run();
        }
    }
}