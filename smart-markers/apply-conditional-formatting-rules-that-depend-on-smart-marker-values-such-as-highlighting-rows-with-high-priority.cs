// Title: How to apply conditional formatting to smart‑marker‑generated rows based on Priority values using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a workbook with smart markers for a task list, processes the markers with WorkbookDesigner, and adds an expression‑based conditional formatting rule that colors rows yellow when the Priority column contains "High". | Modify the conditional formatting to target rows where Priority equals "Low" and set the background color to LightGray, ensuring the rule automatically adapts to the data range produced by the smart markers. | Demonstrate adding two separate expression conditions to the same smart‑marker range—one for "High" priority (yellow) and another for "Medium" priority (orange)—and configure StopIfTrue for each rule.
// Common Searches: asp.net apply conditional formatting to smart marker rows based on priority Aspose.Cells | c# Aspose.Cells conditional formatting expression for dynamic range after WorkbookDesigner.Process | highlight high priority tasks in Excel using smart markers and conditional formatting | how to set background color for rows where column B equals 'High' with Aspose.Cells C#
// Tags: Aspose.Cells conditional formatting with smart markers | C# expression‑based conditional formatting for Excel | highlight rows by priority column using Aspose.Cells | dynamic range handling after WorkbookDesigner processing | set background color for high priority tasks in .xlsx

using System;
using System.Collections.Generic;
using System.Drawing;
using Aspose.Cells;

// Alias to avoid ambiguity with System.Range.
using CellsRange = Aspose.Cells.Range;

namespace AsposeCellsSmartMarkerConditionalFormatting
{
    // Simple data class representing a task with a priority.
    // // Creates a workbook, inserts smart markers for a task list, processes them with WorkbookDesigner, and applies an expression‑based conditional formatting rule that fills rows yellow when the Priority column equals "High".
    public class TaskItem
    {
        public string TaskName { get; set; } = null!;
        public string Priority { get; set; } = null!;
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add header cells.
                sheet.Cells["A1"].PutValue("Task");
                sheet.Cells["B1"].PutValue("Priority");

                // Insert smart markers for the data rows.
                sheet.Cells["A2"].PutValue("&=Tasks.TaskName");
                sheet.Cells["B2"].PutValue("&=Tasks.Priority");

                // Define the smart marker range (required for processing).
                CellsRange smRange = sheet.Cells.CreateRange("A2:B2");
                smRange.Name = "_CellsSmartMarkers";

                // Prepare sample data.
                List<TaskItem> tasks = new List<TaskItem>
                {
                    new TaskItem { TaskName = "Design UI", Priority = "High" },
                    new TaskItem { TaskName = "Write Docs", Priority = "Low" },
                    new TaskItem { TaskName = "Implement Feature", Priority = "Medium" },
                    new TaskItem { TaskName = "Code Review", Priority = "High" }
                };

                // Set up the designer, bind the data source, and process the smart markers.
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };
                designer.SetDataSource("Tasks", tasks);
                designer.Process();

                // Determine the data range after processing.
                int lastDataRow = sheet.Cells.MaxDataRow; // includes header row.

                // Apply conditional formatting to highlight rows where Priority = "High".
                int cfIndex = sheet.ConditionalFormattings.Add();
                FormatConditionCollection fcs = sheet.ConditionalFormattings[cfIndex];

                // Define the area covering the data rows (excluding header).
                CellArea dataArea = new CellArea
                {
                    StartRow = 1,               // Row 2 in Excel (zero‑based index)
                    EndRow = lastDataRow,
                    StartColumn = 0,
                    EndColumn = 1               // Columns A and B
                };
                fcs.AddArea(dataArea);

                // Add an expression condition: =$B2="High"
                int condIdx = fcs.AddCondition(FormatConditionType.Expression);
                FormatCondition fc = fcs[condIdx];
                fc.Formula1 = "=$B2=\"High\"";
                fc.Style.BackgroundColor = Color.Yellow;
                fc.StopIfTrue = true; // Prevent lower‑priority rules from overriding.

                // Save the resulting workbook.
                string outputPath = "SmartMarkerConditionalFormatting.xlsx";
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
