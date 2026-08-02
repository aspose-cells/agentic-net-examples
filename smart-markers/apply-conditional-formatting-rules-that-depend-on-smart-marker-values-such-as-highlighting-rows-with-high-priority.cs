// Title: C# – Apply Conditional Formatting to Smart‑Marker Generated Rows (Highlight High‑Priority) with Aspose.Cells
// Description: This example creates a workbook, inserts smart markers for Task and Priority columns, feeds a dynamic list to WorkbookDesigner, processes the markers, calculates the populated range, and adds an expression‑based conditional formatting rule that colors rows LightSalmon when the Priority column equals "High". The file is saved as an XLSX document.
// Keywords: Aspose.Cells | C# | smart markers | conditional formatting | highlight high priority rows | WorkbookDesigner.Process | dynamic data source | expression condition | CellArea range | Excel automation
// Common Searches: Aspose.Cells add conditional formatting after smart marker processing | C# highlight rows where smart marker column equals High | How to use expression condition with smart markers in Aspose.Cells | Set background color for smart‑marker generated rows | Conditional formatting for dynamic data in Aspose.Cells
// Developer Intent: Generate a worksheet from a collection using smart markers and automatically apply conditional formatting that emphasizes rows with a Priority of "High".
// Use Cases: Create a task‑list export where high‑priority items are visually flagged. | Build a reporting template that colors rows meeting a specific smart‑marker condition. | Automate Excel generation from code‑first data models with built‑in visual cues.
// AI Prompts: Write C# code using Aspose.Cells to add conditional formatting to a smart‑marker range based on a column value. | Show how to define a CellArea that covers all rows produced by smart markers and apply an expression like =$B2="High" to color those rows. | Explain the steps to ensure conditional formatting runs after WorkbookDesigner.Process() in Aspose.Cells.

using System;
using System.Collections.Generic;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// This example creates a workbook, inserts smart markers for Task and Priority columns, feeds a dynamic list to WorkbookDesigner, processes the markers, calculates the populated range, and adds an expression‑based conditional formatting rule that colors rows LightSalmon when the Priority column equals "High". The file is saved as an XLSX document.
class SmartMarkerConditionalFormatting
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add column headers
            sheet.Cells["A1"].PutValue("Task");
            sheet.Cells["B1"].PutValue("Priority");

            // Insert smart markers for the data rows
            sheet.Cells["A2"].PutValue("&=Data.Task");
            sheet.Cells["B2"].PutValue("&=Data.Priority");

            // Define the smart marker range (required for processing)
            AsposeRange smRange = sheet.Cells.CreateRange("A2:B2");
            smRange.Name = "_CellsSmartMarkers";

            // Prepare a data source with tasks and priorities
            var items = new List<dynamic>
            {
                new { Task = "Design UI", Priority = "High" },
                new { Task = "Write Docs", Priority = "Low" },
                new { Task = "Implement Feature", Priority = "High" },
                new { Task = "Testing", Priority = "Medium" }
            };

            // Set the data source and process the smart markers
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };
            designer.SetDataSource("Data", items);
            designer.Process();

            // Determine the last row that contains data after processing
            int lastRow = sheet.Cells.MaxDataRow;

            // Add conditional formatting to highlight rows where Priority = "High"
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];

            // Define the area covering the populated data (columns A and B)
            CellArea area = new CellArea
            {
                StartRow = 1,          // Row 2 in Excel (zero‑based index)
                EndRow = lastRow,
                StartColumn = 0,       // Column A
                EndColumn = 1          // Column B
            };
            fcc.AddArea(area);

            // Add an expression condition: =$B2="High"
            int condIdx = fcc.AddCondition(
                FormatConditionType.Expression,
                OperatorType.None,
                "=$B2=\"High\"",
                null);
            FormatCondition fc = fcc[condIdx];
            fc.Style.BackgroundColor = Color.LightSalmon;
            fc.StopIfTrue = true; // Prevent lower‑priority rules from overriding

            // Save the workbook
            workbook.Save("SmartMarkerConditionalFormatting.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
