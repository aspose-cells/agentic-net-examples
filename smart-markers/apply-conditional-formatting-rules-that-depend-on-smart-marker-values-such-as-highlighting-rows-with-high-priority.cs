using System;
using System.Collections.Generic;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerConditionalFormatting
{
    // Sample data class
    public class TaskItem
    {
        public string Name { get; set; }
        public string Priority { get; set; }   // Expected values: "High", "Medium", "Low"
        public DateTime DueDate { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // 2. Set up header row
            sheet.Cells["A1"].PutValue("Task");
            sheet.Cells["B1"].PutValue("Priority");
            sheet.Cells["C1"].PutValue("Due Date");

            // 3. Insert smart markers for data rows (starting at row 2)
            //    &=$Name, &=$Priority, &=$DueDate are the markers that will be replaced by the data source
            sheet.Cells["A2"].PutValue("&=$Name");
            sheet.Cells["B2"].PutValue("&=$Priority");
            sheet.Cells["C2"].PutValue("&=$DueDate");

            // 4. Prepare sample data source
            List<TaskItem> tasks = new List<TaskItem>
            {
                new TaskItem { Name = "Design UI", Priority = "High", DueDate = DateTime.Today.AddDays(2) },
                new TaskItem { Name = "Write Docs", Priority = "Medium", DueDate = DateTime.Today.AddDays(5) },
                new TaskItem { Name = "Code Review", Priority = "Low", DueDate = DateTime.Today.AddDays(1) },
                new TaskItem { Name = "Deploy", Priority = "High", DueDate = DateTime.Today.AddDays(7) }
            };

            // 5. Configure WorkbookDesigner with the data source and process smart markers
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource("Data", tasks);
            designer.Process();   // lifecycle process (smart markers are populated)

            // 6. After processing, apply conditional formatting to highlight rows where Priority = "High"
            //    Determine the data range (rows 2 to 2 + tasks.Count - 1, columns A:C)
            int startRow = 1; // zero‑based index for row 2
            int endRow = startRow + tasks.Count - 1;
            CellArea dataArea = new CellArea
            {
                StartRow = startRow,
                EndRow = endRow,
                StartColumn = 0,   // column A
                EndColumn = 2      // column C
            };

            // Add a new ConditionalFormatting collection
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];
            fcc.AddArea(dataArea);

            // Add a condition that checks the value in column B (Priority) equals "High"
            // Use CellValue type with OperatorType.Equal and formula referencing the cell itself.
            // Formula1: =B2 (relative reference, will adjust per row)
            // Formula2: "High"
            int conditionIdx = fcc.AddCondition(FormatConditionType.CellValue, OperatorType.Equal, "=B2", "\"High\"");
            FormatCondition condition = fcc[conditionIdx];

            // Set the style to highlight the entire row (background color)
            condition.Style.BackgroundColor = Color.LightCoral;
            condition.Style.Font.Color = Color.White;
            condition.Style.Font.IsBold = true;

            // Ensure this rule has the highest priority
            condition.Priority = 1;
            condition.StopIfTrue = true;

            // 7. Save the workbook (lifecycle save)
            workbook.Save("SmartMarkerConditionalFormatting.xlsx");
        }
    }
}