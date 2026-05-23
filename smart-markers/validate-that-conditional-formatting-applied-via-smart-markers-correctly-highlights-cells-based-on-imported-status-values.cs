using System;
using System.Data;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace SmartMarkerConditionalFormattingValidation
{
    class Program
    {
        static void Main()
        {
            try
            {
                // -------------------- Create template workbook --------------------
                Workbook template = new Workbook();
                Worksheet sheet = template.Worksheets[0];
                Cells cells = sheet.Cells;

                // Header
                cells["A1"].PutValue("Status");

                // Smart marker for status values (line‑by‑line processing)
                cells["A2"].PutValue("&=$Status");

                // Define the range that will receive data (including header)
                // The range name "_CellsSmartMarkers" tells the designer to process this block
                Aspose.Cells.Range dataRange = cells.CreateRange("A1:A2");
                dataRange.Name = "_CellsSmartMarkers";

                // -------------------- Add conditional formatting --------------------
                // Apply formatting to the whole column (A2:A100) – enough rows for data
                int cfIndex = sheet.ConditionalFormattings.Add();
                FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];

                // Define the area for conditional formatting
                CellArea area = new CellArea
                {
                    StartRow = 1,   // row 2 (zero‑based)
                    EndRow = 99,    // row 100
                    StartColumn = 0,
                    EndColumn = 0
                };
                fcc.AddArea(area);

                // Approved → Green background
                int condApproved = fcc.AddCondition(FormatConditionType.ContainsText);
                FormatCondition fcApproved = fcc[condApproved];
                fcApproved.Text = "Approved";
                fcApproved.Style.BackgroundColor = Color.LightGreen;

                // Pending → Yellow background
                int condPending = fcc.AddCondition(FormatConditionType.ContainsText);
                FormatCondition fcPending = fcc[condPending];
                fcPending.Text = "Pending";
                fcPending.Style.BackgroundColor = Color.LightYellow;

                // Rejected → Red background
                int condRejected = fcc.AddCondition(FormatConditionType.ContainsText);
                FormatCondition fcRejected = fcc[condRejected];
                fcRejected.Text = "Rejected";
                fcRejected.Style.BackgroundColor = Color.LightCoral;

                // -------------------- Prepare data source --------------------
                DataTable dt = new DataTable("StatusTable");
                dt.Columns.Add("Status", typeof(string));
                dt.Rows.Add("Approved");
                dt.Rows.Add("Pending");
                dt.Rows.Add("Rejected");
                dt.Rows.Add("Approved");
                dt.Rows.Add("Pending");

                // -------------------- Process smart markers --------------------
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = template
                    // LineByLine is obsolete; range smart markers are used automatically
                };
                designer.SetDataSource(dt);
                designer.Process();

                // -------------------- Validate conditional formatting results --------------------
                Console.WriteLine("Validation of conditional formatting after smart marker processing:");
                for (int row = 1; row <= dt.Rows.Count; row++) // rows start at index 1 (A2)
                {
                    Cell cell = sheet.Cells[row, 0]; // Column A
                    string status = cell.StringValue;
                    ConditionalFormattingResult result = cell.GetConditionalFormattingResult();

                    bool isHighlighted = result?.ConditionalStyle != null;
                    string highlightInfo = isHighlighted
                        ? $"Highlighted (BG: {result.ConditionalStyle.BackgroundColor.Name})"
                        : "Not highlighted";

                    Console.WriteLine($"Row {row + 1}: Status = \"{status}\" -> {highlightInfo}");
                }

                // -------------------- Save the workbook (optional) --------------------
                // Example of safe file saving with existence check
                string outputPath = "SmartMarkerConditionalFormattingResult.xlsx";
                try
                {
                    designer.Workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                // General exception handling to prevent crashes
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}