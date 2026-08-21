// Title: Highlight High‑Priority Rows with Conditional Formatting after Smart Marker Expansion – Aspose.Cells C# Example
// Description: Demonstrates how to load a template workbook containing smart markers, bind a DataTable with Task and Priority fields, process the markers, and apply a ContainsText conditional format that colors cells with "High" priority in light coral with bold white text. The workbook is then saved with the formatting applied.
// Keywords: Aspose.Cells | C# | .NET | smart markers | conditional formatting | ContainsText condition | highlight high priority | Excel export | data‑driven styling | template workbook
// Common Searches: Aspose.Cells conditional formatting after smart markers | C# highlight rows with high priority in Excel | How to use ContainsText in Aspose.Cells | Apply formatting to smart marker generated rows | Excel template with smart markers and conditional colors
// Developer Intent: Add a conditional formatting rule that colors rows marked as "High" priority after smart marker processing.
// Use Cases: Automatically color‑code task lists generated from a database, emphasizing urgent items. | Create reporting templates where smart markers fill data and critical rows stand out visually. | Export Excel sheets from .NET applications with built‑in styling for status‑based values.
// AI Prompts: Modify the example to also format rows where Priority equals "Medium" with a yellow background. | Generate C# code that applies a ContainsText conditional format to multiple columns after smart marker expansion. | Explain how to replace the ContainsText rule with a formula‑based condition that highlights rows where Priority = "High".

using System;
using System.Data;
using Aspose.Cells;
using System.Drawing;

namespace AsposeCellsConditionalFormattingSmartMarkers
{
    // Demonstrates how to load a template workbook containing smart markers, bind a DataTable with Task and Priority fields, process the markers, and apply a ContainsText conditional format that colors cells with "High" priority in light coral with bold white text. The workbook is then saved with the formatting applied.
    class Program
    {
        static void Main()
        {
            // Load the template workbook that contains smart markers (e.g., &amp;=Tasks.Task, &amp;=Tasks.Priority)
            Workbook workbook = new Workbook("Template.xlsx");

            // Initialize the WorkbookDesigner with the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Prepare a data source with a Priority column
            DataTable tasksTable = new DataTable("Tasks");
            tasksTable.Columns.Add("Task", typeof(string));
            tasksTable.Columns.Add("Priority", typeof(string));

            // Sample data rows
            tasksTable.Rows.Add("Prepare report", "High");
            tasksTable.Rows.Add("Team meeting", "Medium");
            tasksTable.Rows.Add("Code review", "Low");
            tasksTable.Rows.Add("Client call", "High");
            tasksTable.Rows.Add("Documentation", "Medium");

            // Set the data source for the smart markers
            designer.SetDataSource(tasksTable);

            // Process the smart markers – this expands the rows based on the data source
            designer.Process();

            // After processing, apply conditional formatting to highlight rows with high priority
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the range that contains the populated data (excluding header row)
            int firstDataRow = 1; // assuming row 0 is header
            int lastDataRow = sheet.Cells.MaxDataRow;
            int priorityColumnIndex = 1; // column B (0‑based index)

            // Add a new conditional formatting collection to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcs = sheet.ConditionalFormattings[cfIndex];

            // Define the area to which the condition will be applied (entire Priority column)
            CellArea priorityArea = new CellArea
            {
                StartRow = firstDataRow,
                EndRow = lastDataRow,
                StartColumn = priorityColumnIndex,
                EndColumn = priorityColumnIndex
            };
            fcs.AddArea(priorityArea);

            // Add a condition that looks for the text "High" in the Priority column
            int conditionIndex = fcs.AddCondition(FormatConditionType.ContainsText);
            FormatCondition condition = fcs[conditionIndex];
            condition.Text = "High";

            // Set the visual style for rows that meet the condition (e.g., light red background)
            condition.Style.BackgroundColor = Color.LightCoral;
            condition.Style.Font.Color = Color.White;
            condition.Style.Font.IsBold = true;

            // Save the resulting workbook
            workbook.Save("Output_WithConditionalFormatting.xlsx");
        }
    }
}
