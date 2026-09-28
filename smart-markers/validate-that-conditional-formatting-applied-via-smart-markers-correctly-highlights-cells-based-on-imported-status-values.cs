// Title: Validate smart‑marker conditional formatting for task status cells using Aspose.Cells for .NET
// AI Prompts: Create a C# workbook template with smart markers for Task and Status, add text‑contains conditional formatting rules for "Completed" (green) and "Pending" (yellow), then process a DataTable line‑by‑line and verify the applied colors. | Write code that uses WorkbookDesigner to populate smart‑marker rows, defines conditional formatting on column B, retrieves ConditionalFormattingResult for each status cell, and outputs pass/fail based on expected background colors. | Generate an Excel file named SmartMarkerConditionalFormattingResult.xlsx that demonstrates conditional formatting validation after smart‑marker processing, printing validation results to the console.
// Common Searches: asp.net how to test conditional formatting applied by smart markers in Aspose.Cells | c# verify conditional formatting colors after WorkbookDesigner.Process | example of line‑by‑line smart markers with conditional formatting in Aspose.Cells | check if text‑contains conditional formatting works for imported data in Excel using Aspose.Cells
// Tags: smart markers conditional formatting Aspose.Cells | WorkbookDesigner line-by-line data import C# | validate conditional formatting result Aspose.Cells | Excel conditional formatting based on cell text C# | retrieve ConditionalFormattingResult programmatically

using System;
using System.Data;
using System.Drawing;
using Aspose.Cells;

namespace SmartMarkerConditionalFormattingValidation
{
    // The example builds a template workbook, inserts smart markers for task names and statuses, defines two text‑contains conditional formatting rules on the Status column (green for "Completed", yellow for "Pending"), populates a DataTable with sample tasks, processes the smart markers line‑by‑line using WorkbookDesigner, iterates over the resulting rows to retrieve each cell's ConditionalFormattingResult, compares the actual background color with the expected one, prints pass/fail messages for each row, and saves the final workbook as SmartMarkerConditionalFormattingResult.xlsx.
    class Program
    {
        static void Main()
        {
            // ------------------- Create template workbook -------------------
            Workbook template = new Workbook();
            Worksheet sheet = template.Worksheets[0];
            Cells cells = sheet.Cells;

            // Header row
            cells["A1"].PutValue("Task");
            cells["B1"].PutValue("Status");

            // Smart markers for data rows (line‑by‑line processing)
            cells["A2"].PutValue("&=Tasks.TaskName");
            cells["B2"].PutValue("&=Tasks.Status");

            // Define conditional formatting on the Status column (B)
            // Highlight "Completed" with green background
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];

            // Apply to a large range that will cover all imported rows
            CellArea statusArea = new CellArea
            {
                StartRow = 1,   // row 2 (zero‑based)
                EndRow = 100,   // enough rows for data
                StartColumn = 1,
                EndColumn = 1
            };
            fcc.AddArea(statusArea);

            // Condition 1: ContainsText "Completed"
            int condIdx1 = fcc.AddCondition(FormatConditionType.ContainsText);
            FormatCondition cond1 = fcc[condIdx1];
            cond1.Text = "Completed";
            cond1.Style.BackgroundColor = Color.LightGreen;

            // Condition 2: ContainsText "Pending"
            int condIdx2 = fcc.AddCondition(FormatConditionType.ContainsText);
            FormatCondition cond2 = fcc[condIdx2];
            cond2.Text = "Pending";
            cond2.Style.BackgroundColor = Color.LightYellow;

            // ------------------- Prepare data source -------------------
            DataTable dt = new DataTable("Tasks");
            dt.Columns.Add("TaskName", typeof(string));
            dt.Columns.Add("Status", typeof(string));

            dt.Rows.Add("Design UI", "Completed");
            dt.Rows.Add("Implement Backend", "Pending");
            dt.Rows.Add("Write Tests", "Completed");
            dt.Rows.Add("Deploy", "Pending");

            // ------------------- Process smart markers -------------------
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = template,
                LineByLine = true   // process line by line (default)
            };
            designer.SetDataSource(dt);
            designer.Process();

            // ------------------- Validation of conditional formatting -------------------
            // Expected background colors per status
            Color expectedCompleted = Color.LightGreen;
            Color expectedPending = Color.LightYellow;

            // Iterate over data rows (starting at row index 1)
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                int rowIndex = i + 1; // because data starts at row 2 in the sheet
                Cell statusCell = cells[rowIndex, 1]; // column B

                // Retrieve the conditional formatting result for the cell
                ConditionalFormattingResult result = statusCell.GetConditionalFormattingResult();

                // Determine expected color based on the source data
                string statusValue = dt.Rows[i]["Status"].ToString();
                Color expectedColor = statusValue == "Completed" ? expectedCompleted : expectedPending;

                // Check if the conditional style was applied and matches expectation
                bool isCorrect = result?.ConditionalStyle != null &&
                                 result.ConditionalStyle.BackgroundColor.Equals(expectedColor);

                Console.WriteLine($"Row {rowIndex + 1}: Status=\"{statusValue}\" " +
                                  $"=> Conditional formatting {(isCorrect ? "PASSED" : "FAILED")}");
            }

            // ------------------- Save the result workbook -------------------
            designer.Workbook.Save("SmartMarkerConditionalFormattingResult.xlsx");
        }
    }
}
