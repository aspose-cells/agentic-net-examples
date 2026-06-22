using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace DuplicateHighlightDemo
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
                Cells cells = sheet.Cells;

                // Populate sample data with some duplicate values
                // Column A will contain: Dup, Unique1, Dup, Unique3, Dup
                for (int i = 0; i < 5; i++)
                {
                    string value = (i % 2 == 0) ? "Dup" : $"Unique{i}";
                    cells[i, 0].PutValue(value);
                }

                // Create a range covering the data and assign a name to it
                AsposeRange dataRange = cells.CreateRange("A1", "A5");
                dataRange.Name = "DataRange";

                // Add a conditional formatting collection to the worksheet
                int cfIndex = sheet.ConditionalFormattings.Add();
                FormatConditionCollection fcs = sheet.ConditionalFormattings[cfIndex];

                // Define the area of the conditional formatting (the named range)
                CellArea area = new CellArea
                {
                    StartRow = dataRange.FirstRow,
                    EndRow = dataRange.FirstRow + dataRange.RowCount - 1,
                    StartColumn = dataRange.FirstColumn,
                    EndColumn = dataRange.FirstColumn + dataRange.ColumnCount - 1
                };
                fcs.AddArea(area);

                // Add a condition that highlights duplicate values
                int conditionIndex = fcs.AddCondition(FormatConditionType.DuplicateValues);
                FormatCondition duplicateCondition = fcs[conditionIndex];

                // Create a style for highlighting (yellow background)
                Style highlightStyle = workbook.CreateStyle();
                highlightStyle.ForegroundColor = Color.Yellow;
                highlightStyle.Pattern = BackgroundType.Solid;
                duplicateCondition.Style = highlightStyle;

                // Save the workbook
                string outputPath = "DuplicateHighlight.xlsx";
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