// Title: C# – Validate Conditional Formatting Applied via Smart Markers in Aspose.Cells
// Description: Creates an in‑memory template with smart markers for Item and Status, fills it using a List<DataItem> through WorkbookDesigner, adds text‑based conditional formatting (Completed → LightGreen, Pending → LightYellow, Failed → LightCoral) to column B, then reads each cell's ConditionalFormattingResult to confirm the background color matches the status value before saving the workbook.
// Keywords: Aspose.Cells | C# | smart markers | conditional formatting validation | status column coloring | WorkbookDesigner | programmatic color check | Excel automation testing
// Common Searches: aspnet verify conditional formatting after smart markers | aspnet check cell background color Aspose.Cells | unit test conditional formatting Aspose.Cells | how to read ConditionalFormattingResult C#
// Developer Intent: Confirm that conditional formatting rules added after processing smart markers correctly highlight Status cells according to their values.
// Use Cases: Generate a status‑driven report where rows are automatically colored and programmatically verified. | Automate regression tests for conditional formatting in workbooks built with smart markers. | Create a reusable template that applies and validates visual cues for Completed, Pending, and Failed items.
// AI Prompts: Write C# code that adds text‑based conditional formatting for 'Completed', 'Pending', and 'Failed' after processing smart markers with Aspose.Cells and verifies the applied colors. | Show how to retrieve ConditionalFormattingResult for a cell and compare its background color to an expected value in Aspose.Cells .NET. | Explain how to build a unit test that asserts conditional formatting matches data values in a workbook generated via smart markers.

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// Creates an in‑memory template with smart markers for Item and Status, fills it using a List<DataItem> through WorkbookDesigner, adds text‑based conditional formatting (Completed → LightGreen, Pending → LightYellow, Failed → LightCoral) to column B, then reads each cell's ConditionalFormattingResult to confirm the background color matches the status value before saving the workbook.
class Program
{
    static void Main()
    {
        // 1. Create a template workbook with smart markers
        Workbook template = new Workbook();
        Worksheet ws = template.Worksheets[0];
        Cells cells = ws.Cells;

        // Header row
        cells["A1"].PutValue("Item");
        cells["B1"].PutValue("Status");

        // Smart markers for data rows (line‑by‑line processing)
        cells["A2"].PutValue("&=Data.Item");
        cells["B2"].PutValue("&=Data.Status");

        // Save the template to a memory stream (create rule)
        using (MemoryStream ms = new MemoryStream())
        {
            template.Save(ms, SaveFormat.Xlsx);
            ms.Position = 0;

            // 2. Load the template into WorkbookDesigner (load rule)
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = new Workbook(ms);

            // Optional: display detected smart markers
            string[] markers = designer.GetSmartMarkers();
            Console.WriteLine("Smart markers found:");
            foreach (string m in markers) Console.WriteLine(m);

            // 3. Prepare data source
            List<DataItem> data = new List<DataItem>
            {
                new DataItem { Item = "Task1", Status = "Completed" },
                new DataItem { Item = "Task2", Status = "Pending" },
                new DataItem { Item = "Task3", Status = "Failed" }
            };
            designer.SetDataSource("Data", data);

            // 4. Process smart markers (populate data)
            designer.Process();

            // 5. Apply conditional formatting to the Status column (B)
            Worksheet resultSheet = designer.Workbook.Worksheets[0];
            // Define a range that covers possible rows (B2:B100)
            CellArea statusArea = new CellArea { StartRow = 1, EndRow = 100, StartColumn = 1, EndColumn = 1 };
            int cfIndex = resultSheet.ConditionalFormattings.Add();
            FormatConditionCollection fcc = resultSheet.ConditionalFormattings[cfIndex];
            fcc.AddArea(statusArea);

            // Completed → LightGreen
            int condIdx = fcc.AddCondition(FormatConditionType.ContainsText);
            FormatCondition fc = fcc[condIdx];
            fc.Text = "Completed";
            fc.Style.BackgroundColor = Color.LightGreen;

            // Pending → LightYellow
            condIdx = fcc.AddCondition(FormatConditionType.ContainsText);
            fc = fcc[condIdx];
            fc.Text = "Pending";
            fc.Style.BackgroundColor = Color.LightYellow;

            // Failed → LightCoral
            condIdx = fcc.AddCondition(FormatConditionType.ContainsText);
            fc = fcc[condIdx];
            fc.Text = "Failed";
            fc.Style.BackgroundColor = Color.LightCoral;

            // 6. Validate that conditional formatting highlights cells correctly
            Console.WriteLine("\nValidation Results:");
            for (int i = 0; i < data.Count; i++)
            {
                int rowIndex = i + 1; // zero‑based index (row 1 = second row in sheet)
                Cell statusCell = resultSheet.Cells[rowIndex, 1]; // Column B
                ConditionalFormattingResult cfResult = statusCell.GetConditionalFormattingResult();

                string expectedStatus = data[i].Status;
                Color expectedColor = expectedStatus == "Completed" ? Color.LightGreen :
                                      expectedStatus == "Pending"   ? Color.LightYellow :
                                      expectedStatus == "Failed"    ? Color.LightCoral :
                                      Color.Empty;

                bool isMatch = cfResult != null &&
                               cfResult.ConditionalStyle != null &&
                               cfResult.ConditionalStyle.BackgroundColor.Equals(expectedColor);

                Console.WriteLine($"Row {rowIndex + 1}: Status='{expectedStatus}' " +
                                  $"=> Highlighted={(cfResult?.ConditionalStyle != null)} " +
                                  $"Match={isMatch}");
            }

            // 7. Save the final workbook (save rule)
            designer.Workbook.Save("ConditionalFormattingValidation.xlsx");
        }
    }
}

// Simple POCO for data source
public class DataItem
{
    public string Item { get; set; }
    public string Status { get; set; }
}
