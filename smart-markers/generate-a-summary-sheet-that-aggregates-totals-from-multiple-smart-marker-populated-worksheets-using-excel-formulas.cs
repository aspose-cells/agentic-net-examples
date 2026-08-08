// Title: C# – Add a Summary Sheet that Totals a Column Across Smart‑Marker Worksheets with Aspose.Cells
// Description: This example loads a template workbook containing Smart Markers, binds SalesData and Inventory DataTables, processes the markers, creates a "Summary" worksheet, and automatically writes a SUM formula for column C of each source sheet. The formulas are calculated with CalculateFormula and the workbook is saved as OutputWithSummary.xlsx.
// Keywords: Aspose.Cells | C# | Smart Marker | summary worksheet | column sum | Excel formula generation | WorkbookDesigner | multiple worksheets aggregation | CalculateFormula | dynamic range reference | GitHub example
// Common Searches: Aspose.Cells create summary sheet after smart marker processing | C# sum column across worksheets using Aspose.Cells | How to add SUM formulas with Aspose.Cells WorkbookDesigner | Generate totals sheet for smart marker data in .NET | Calculate formulas programmatically in Aspose.Cells
// Developer Intent: Automatically generate a summary sheet that lists each source worksheet and the summed values of a specific column using Excel formulas.
// Use Cases: After populating several worksheets via Smart Markers, produce a consolidated "Summary" sheet that shows the total Amount or Value from column C of each sheet. | Create dynamic SUM formulas that adapt to the actual data range on each worksheet, so the totals update when source data changes. | Trigger CalculateFormula to materialize the totals before saving the workbook for downstream reporting.
// AI Prompts: Write C# code with Aspose.Cells to add a summary worksheet that iterates over all existing sheets, builds a SUM formula for column C of each sheet, and writes the sheet name and formula into the summary. | Show how to escape worksheet names containing special characters when constructing Excel formulas in Aspose.Cells and how to invoke CalculateFormula after adding the summary. | Explain how to modify the example to aggregate a different numeric column or to include multiple aggregated columns in the summary sheet.

using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerSummary
{
    // This example loads a template workbook containing Smart Markers, binds SalesData and Inventory DataTables, processes the markers, creates a "Summary" worksheet, and automatically writes a SUM formula for column C of each source sheet. The formulas are calculated with CalculateFormula and the workbook is saved as OutputWithSummary.xlsx.
    class Program
    {
        static void Main()
        {
            // Load the template workbook that contains smart markers
            Workbook template = new Workbook("TemplateWithSmartMarkers.xlsx");

            // Prepare sample data sources for multiple worksheets
            // Assume the template has smart markers for tables named "SalesData" and "Inventory"
            DataTable salesTable = new DataTable("SalesData");
            salesTable.Columns.Add("Product", typeof(string));
            salesTable.Columns.Add("Quantity", typeof(int));
            salesTable.Columns.Add("Amount", typeof(double));
            salesTable.Rows.Add("A", 10, 250.0);
            salesTable.Rows.Add("B", 5, 150.0);
            salesTable.Rows.Add("C", 8, 200.0);

            DataTable inventoryTable = new DataTable("Inventory");
            inventoryTable.Columns.Add("Item", typeof(string));
            inventoryTable.Columns.Add("Stock", typeof(int));
            inventoryTable.Columns.Add("Value", typeof(double));
            inventoryTable.Rows.Add("X", 20, 500.0);
            inventoryTable.Rows.Add("Y", 15, 375.0);
            inventoryTable.Rows.Add("Z", 30, 750.0);

            // Initialize WorkbookDesigner with the template workbook
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = template
            };

            // Bind the data sources to the corresponding smart marker names
            designer.SetDataSource("SalesData", salesTable);
            designer.SetDataSource("Inventory", inventoryTable);

            // Process all smart markers in the workbook
            designer.Process();

            // After processing, create a summary worksheet
            Worksheet summarySheet = designer.Workbook.Worksheets.Add("Summary");

            // Header row for the summary sheet
            summarySheet.Cells["A1"].PutValue("Source Sheet");
            summarySheet.Cells["B1"].PutValue("Total Amount/Value");

            // Row index where we start writing summary data
            int summaryRow = 1; // zero‑based index (row 2 in Excel)

            // Iterate through all worksheets except the newly added summary sheet
            for (int i = 0; i < designer.Workbook.Worksheets.Count; i++)
            {
                Worksheet ws = designer.Workbook.Worksheets[i];
                if (ws.Name == "Summary")
                    continue; // skip the summary sheet itself

                // Determine the last row that contains data in the worksheet
                int lastDataRow = ws.Cells.MaxDataRow; // zero‑based index

                // Assume the numeric column to aggregate is column C (index 2)
                // Build the SUM formula referencing the current worksheet
                string sheetName = ws.Name.Replace("'", "''"); // escape single quotes
                string formula = $"=SUM('{sheetName}'!C2:C{lastDataRow + 1})";

                // Write sheet name and formula into the summary sheet
                summarySheet.Cells[summaryRow, 0].PutValue(ws.Name);          // Column A
                summarySheet.Cells[summaryRow, 1].Formula = formula;        // Column B

                summaryRow++;
            }

            // Optionally calculate formulas so that the totals are materialized
            designer.Workbook.CalculateFormula();

            // Save the final workbook with the summary sheet
            designer.Workbook.Save("OutputWithSummary.xlsx");
        }
    }
}
