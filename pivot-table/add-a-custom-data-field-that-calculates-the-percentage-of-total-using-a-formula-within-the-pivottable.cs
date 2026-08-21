// Title: C# – Add a Percentage‑of‑Total Calculated Field to an Aspose.Cells PivotTable
// Description: This example creates a workbook, inserts sample Category and Sales data, builds a PivotTable, adds a row field (Category) and a data field (Sales), then uses AddCalculatedField (without a leading '=') to create a "PctOfTotal" field that computes Sales/Total. The field is formatted as a percentage, the PivotTable is refreshed and calculated, and the file is saved as PivotTable_PercentageOfTotal.xlsx.
// Keywords: Aspose.Cells C# PivotTable | AddCalculatedField percentage of total | PivotTable custom formula Aspose | format pivot field as percentage | Aspose.Cells calculated field without equals sign | Excel PivotTable percentage of grand total | C# Aspose.Cells example
// Common Searches: Aspose.Cells add calculated field percentage of total | C# PivotTable percentage of grand total Aspose | AddCalculatedField formula without '=' Aspose.Cells | format pivot data field as % in Aspose.Cells | refresh PivotTable after adding calculated field C#
// Developer Intent: Create a PivotTable in Aspose.Cells and add a custom calculated field that shows each sales value as a percentage of the overall total.
// Use Cases: Financial reports that display sales share by category. | Dashboards where line items are shown as a proportion of total revenue. | Printable Excel summaries with pivot percentages formatted for presentation.
// AI Prompts: Generate C# code with Aspose.Cells to add a calculated field named 'PctOfTotal' that divides the Sales field by the grand total and formats it as a percentage. | Explain why AddCalculatedField in Aspose.Cells requires the formula string without a leading '=' sign. | List the steps to refresh and recalculate a PivotTable after inserting a custom calculated field using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotPercentageExample
{
    // This example creates a workbook, inserts sample Category and Sales data, builds a PivotTable, adds a row field (Category) and a data field (Sales), then uses AddCalculatedField (without a leading '=') to create a "PctOfTotal" field that computes Sales/Total. The field is formatted as a percentage, the PivotTable is refreshed and calculated, and the file is saved as PivotTable_PercentageOfTotal.xlsx.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data
                cells["A1"].Value = "Category";
                cells["B1"].Value = "Sales";
                cells["A2"].Value = "Electronics";
                cells["B2"].Value = 1200;
                cells["A3"].Value = "Electronics";
                cells["B3"].Value = 800;
                cells["A4"].Value = "Furniture";
                cells["B4"].Value = 600;
                cells["A5"].Value = "Furniture";
                cells["B5"].Value = 400;

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "SalesPivot");
                PivotTable pivot = sheet.PivotTables[pivotIndex];

                // Add the row field (Category) and the data field (Sales)
                pivot.AddFieldToArea(PivotFieldType.Row, "Category");
                pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Add a calculated field that computes the percentage of total sales.
                // Note: The formula should not start with '=' when using AddCalculatedField.
                pivot.AddCalculatedField("PctOfTotal", "Sales/Total", true);

                // Retrieve the newly added calculated field and format it as a percentage
                PivotField pctField = pivot.DataFields[pivot.DataFields.Count - 1];
                pctField.NumberFormat = "0.00%";

                // Refresh and calculate the pivot table data
                pivot.RefreshData();
                pivot.CalculateData();

                // Save the workbook
                workbook.Save("PivotTable_PercentageOfTotal.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
