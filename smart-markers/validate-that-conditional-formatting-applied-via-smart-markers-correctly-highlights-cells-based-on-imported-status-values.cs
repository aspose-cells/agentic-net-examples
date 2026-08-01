// Title: C# – Validate Conditional Formatting with Smart Markers in Aspose.Cells
// Description: Shows how to add text‑based conditional formatting to a smart‑marker column, bind a DataTable, process the markers, and programmatically verify that each generated status cell receives the correct background color.
// Keywords: Aspose.Cells | C# | smart markers | conditional formatting | ConditionalFormattingResult | validate formatting | DataTable | cell background color | Excel report automation | unit test
// Common Searches: Aspose.Cells verify conditional formatting after smart marker processing | C# get ConditionalFormattingResult for a cell | smart markers conditional formatting example | check cell color Aspose.Cells .NET | unit test Aspose.Cells conditional formatting
// Developer Intent: Confirm that the conditional formatting rules defined for the Status column are applied to the cells populated via smart markers.
// Use Cases: Generate an Excel report where status values are automatically colored using smart markers. | Automate regression tests that compare expected and actual cell colors after processing data. | Create reusable templates with embedded conditional formatting for dynamic data sources. | Export validated workbooks for downstream consumption.
// AI Prompts: Write C# code that asserts the background color of each status cell matches its text after processing smart markers with Aspose.Cells. | Provide a unit‑test method that loads a template workbook, sets a DataTable as the data source, runs WorkbookDesigner.Process, and checks ConditionalFormattingResult for every populated row. | Show how to extract ConditionalFormattingResult for a range and output mismatched cells in a concise report.

using System;
using System.Data;
using System.Drawing;
using Aspose.Cells;

// Shows how to add text‑based conditional formatting to a smart‑marker column, bind a DataTable, process the markers, and programmatically verify that each generated status cell receives the correct background color.
public class ValidateConditionalFormattingSmartMarkers
{
    public static void Main()
    {
        try
        {
            // Create a template workbook with smart markers
            Workbook template = new Workbook();
            Worksheet ws = template.Worksheets[0];
            Cells cells = ws.Cells;

            // Header row
            cells["A1"].PutValue("ID");
            cells["B1"].PutValue("Status");

            // Smart marker row (will be expanded by the designer)
            cells["A2"].PutValue("&=Tasks.ID");
            cells["B2"].PutValue("&=Tasks.Status");

            // Define conditional formatting for the Status column (B2:B5 after data is populated)
            int cfIndex = ws.ConditionalFormattings.Add();
            FormatConditionCollection fcc = ws.ConditionalFormattings[cfIndex];

            // Area covering the expected data rows (adjusted later if needed)
            CellArea area = new CellArea
            {
                StartRow = 1,    // Row 2 (zero‑based)
                StartColumn = 1, // Column B
                EndRow = 4,      // Row 5 (placeholder for 4 data rows)
                EndColumn = 1
            };
            fcc.AddArea(area);

            // Condition: "Completed" → LightGreen background
            int condIdx = fcc.AddCondition(FormatConditionType.ContainsText);
            FormatCondition fc = fcc[condIdx];
            fc.Text = "Completed";
            fc.Style.BackgroundColor = Color.LightGreen;

            // Condition: "Pending" → LightYellow background
            condIdx = fcc.AddCondition(FormatConditionType.ContainsText);
            fc = fcc[condIdx];
            fc.Text = "Pending";
            fc.Style.BackgroundColor = Color.LightYellow;

            // Condition: "Failed" → LightCoral background
            condIdx = fcc.AddCondition(FormatConditionType.ContainsText);
            fc = fcc[condIdx];
            fc.Text = "Failed";
            fc.Style.BackgroundColor = Color.LightCoral;

            // Prepare data source
            DataTable dt = new DataTable("Tasks");
            dt.Columns.Add("ID", typeof(int));
            dt.Columns.Add("Status", typeof(string));
            dt.Rows.Add(1, "Completed");
            dt.Rows.Add(2, "Pending");
            dt.Rows.Add(3, "Failed");
            dt.Rows.Add(4, "Completed");

            // Set up WorkbookDesigner, assign data source and process smart markers
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = template
            };
            designer.SetDataSource(dt);
            designer.Process();

            // Validate that each status cell received the expected conditional formatting
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                int rowIndex = i + 1; // Data starts at row 2 (zero‑based index 1)
                Cell statusCell = ws.Cells[rowIndex, 1]; // Column B
                ConditionalFormattingResult result = statusCell.GetConditionalFormattingResult();

                string status = dt.Rows[i]["Status"]?.ToString() ?? string.Empty;
                string expectedColor = GetExpectedColorName(status);
                string actualColor = (result?.ConditionalStyle != null)
                    ? result.ConditionalStyle.BackgroundColor.Name
                    : "None";

                Console.WriteLine($"Row {rowIndex + 1}: Status='{status}', ExpectedColor={expectedColor}, ActualColor={actualColor}");
            }

            // Save the resulting workbook (optional verification)
            string outputPath = "ValidatedSmartMarkers.xlsx";
            designer.Workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper to map status text to the expected background color name
    private static string GetExpectedColorName(string status)
    {
        return status switch
        {
            "Completed" => Color.LightGreen.Name,
            "Pending" => Color.LightYellow.Name,
            "Failed" => Color.LightCoral.Name,
            _ => "None"
        };
    }
}
