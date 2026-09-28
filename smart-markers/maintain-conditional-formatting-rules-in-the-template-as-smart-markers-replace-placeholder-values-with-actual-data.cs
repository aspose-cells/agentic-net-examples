// Title: How to retain conditional formatting when filling smart markers in an Aspose.Cells .NET workbook
// AI Prompts: Create an Excel template with smart markers for product and quantity, add a conditional formatting rule that colors cells yellow when the quantity exceeds 50, then run WorkbookDesigner.Process while keeping the formatting intact. | Build a DataTable, bind it to the smart markers using WorkbookDesigner, and generate an output file that still contains the original conditional formatting rules. | Define a named range called _CellsSmartMarkers for line‑by‑line smart‑marker processing, apply a CellValue > 50 condition to column B, and ensure the rule survives after processing the markers.
// Common Searches: Aspose.Cells preserve conditional formatting after WorkbookDesigner.Process with smart markers | C# add conditional formatting to a smart marker range in Excel using Aspose.Cells | how to keep cell style rules when populating smart markers from a DataTable | conditional formatting rule for values greater than 50 with smart markers Aspose.Cells .NET
// Tags: WorkbookDesigner conditional formatting retention | smart markers line‑by‑line range definition | cellvalue greaterthan 50 rule Aspose.Cells | named range _CellsSmartMarkers usage | populate DataTable into smart markers C#

using System;
using System.Data;
using Aspose.Cells;
using System.Drawing;

// The example creates a workbook template, inserts smart markers for product and quantity, defines a named range for line‑by‑line processing, adds a conditional formatting rule that highlights quantities over 50 with a yellow background, binds a DataTable to the markers via WorkbookDesigner, processes the template, and saves the result while retaining the conditional formatting.
class ConditionalFormattingSmartMarkerDemo
{
    static void Main()
    {
        // Create a new workbook that will serve as the template
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add column headers
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Quantity");

        // Insert smart markers that will be replaced by data source values
        sheet.Cells["A2"].PutValue("&=Data.Product");
        sheet.Cells["B2"].PutValue("&=Data.Quantity");

        // Define the range that contains the smart markers.
        // Naming the range as "_CellsSmartMarkers" enables line‑by‑line processing.
        sheet.Cells.CreateRange("A2:B2").Name = "_CellsSmartMarkers";

        // -----------------------------------------------------------------
        // Create a conditional formatting rule that highlights quantities > 50
        // -----------------------------------------------------------------
        int cfIndex = sheet.ConditionalFormattings.Add();                     // Add a new ConditionalFormatting object
        FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex]; // Get its collection

        // Apply the formatting to column B (Quantity column) rows 2‑101
        CellArea area = new CellArea
        {
            StartRow = 1,   // Row index is zero‑based (row 2 in Excel)
            EndRow = 100,
            StartColumn = 1, // Column B
            EndColumn = 1
        };
        fcc.AddArea(area);

        // Add a condition: CellValue GreaterThan 50
        int conditionIdx = fcc.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "50", null);
        FormatCondition condition = fcc[conditionIdx];
        condition.Style.BackgroundColor = Color.Yellow; // Highlight with yellow background

        // -------------------------------------------------
        // Prepare a DataTable that will be bound to the smart markers
        // -------------------------------------------------
        DataTable dt = new DataTable("Data");
        dt.Columns.Add("Product", typeof(string));
        dt.Columns.Add("Quantity", typeof(int));
        dt.Rows.Add("Apple", 30);
        dt.Rows.Add("Banana", 60);
        dt.Rows.Add("Cherry", 45);
        dt.Rows.Add("Date", 80);

        // -------------------------------------------------
        // Set up WorkbookDesigner, bind the data source and process the smart markers
        // -------------------------------------------------
        WorkbookDesigner designer = new WorkbookDesigner(workbook);
        designer.SetDataSource("Data", dt);
        designer.Process(); // Populate the smart markers with data

        // -------------------------------------------------
        // Save the resulting workbook; conditional formatting remains intact
        // -------------------------------------------------
        workbook.Save("ConditionalFormattingSmartMarkerResult.xlsx");
    }
}
