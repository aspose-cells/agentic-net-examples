using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsDataBarExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate column L (index 11) with sample progress percentages
            // For demonstration we fill rows 0 to 9 (L1:L10)
            for (int i = 0; i < 10; i++)
            {
                // Example values: 0%, 10%, ..., 90%
                sheet.Cells[i, 11].PutValue(i * 10);
            }

            // Add an empty conditional formatting collection to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection cfCollection = sheet.ConditionalFormattings[cfIndex];

            // Define the cell area for column L (rows 0-9)
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 9,
                StartColumn = 11,
                EndColumn = 11
            };
            cfCollection.AddArea(area);

            // Add a DataBar condition to the collection
            int conditionIndex = cfCollection.AddCondition(FormatConditionType.DataBar);
            FormatCondition condition = cfCollection[conditionIndex];

            // Configure the DataBar properties
            DataBar dataBar = condition.DataBar;
            dataBar.Color = Color.Green;                                 // Bar color
            dataBar.MinCfvo.Type = FormatConditionValueType.AutomaticMin; // Minimum value
            dataBar.MaxCfvo.Type = FormatConditionValueType.AutomaticMax; // Maximum value
            dataBar.ShowValue = true;                                    // Show cell values alongside bars

            // Save the workbook to a file
            workbook.Save("ColumnL_DataBar.xlsx", SaveFormat.Xlsx);
        }
    }
}