// Title: How to embed per‑row multiplication and grand‑total SUM formulas in Aspose.Cells smart markers using C#
// AI Prompts: Write C# code that adds a multiplication formula referencing the Quantity and Price cells for each smart‑marker row in an Aspose.Cells worksheet. | Generate a C# example that sets a SUM formula for a grand‑total row after WorkbookDesigner processes smart markers in Aspose.Cells.
// Common Searches: C# Aspose.Cells smart marker calculate total column with formula | Aspose.Cells WorkbookDesigner add grand total row after data population | Insert dynamic Excel formulas in smart marker template using C#
// Tags: smart marker per‑row total calculation Aspose.Cells | Aspose.Cells grand total calculation | WorkbookDesigner dynamic formula insertion | C# embed cell formula in smart marker table | Aspose.Cells calculate total column per row

using System;
using System.Data;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example creates a workbook, defines a DataTable of products, inserts smart markers for Name, Quantity, and Price, embeds a per‑row total formula (=B2*C2), adds a grand‑total SUM formula after processing, processes the markers with WorkbookDesigner, and saves the file as SmartMarkerWithFormula.xlsx.
class SmartMarkerFormulaExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Report";

        // Prepare sample data for smart markers
        DataTable dt = new DataTable("Products");
        dt.Columns.Add("Name", typeof(string));
        dt.Columns.Add("Quantity", typeof(int));
        dt.Columns.Add("Price", typeof(double));

        dt.Rows.Add("Apple", 10, 0.5);
        dt.Rows.Add("Banana", 5, 0.3);
        dt.Rows.Add("Orange", 8, 0.4);

        // Insert smart markers into the worksheet
        // Header row
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Quantity");
        sheet.Cells["C1"].PutValue("Price");
        sheet.Cells["D1"].PutValue("Total");

        // Data rows using smart markers
        // ${Products.Name} will be replaced by each product name, etc.
        sheet.Cells["A2"].PutValue("${Products.Name}");
        sheet.Cells["B2"].PutValue("${Products.Quantity}");
        sheet.Cells["C2"].PutValue("${Products.Price}");

        // Embed a formula that references the Quantity and Price cells of the same row
        // The formula will be evaluated after smart marker processing
        // =B2*C2 calculates total price per product
        sheet.Cells["D2"].Formula = "=B2*C2";

        // Add a grand total row below the data
        // The row index will be determined after processing, so we use a placeholder row now
        // The formula will sum the Total column (D) for all data rows
        // We'll set the formula after processing when we know the last data row
        int grandTotalRow = dt.Rows.Count + 2; // Row after the data rows (1-based index)
        sheet.Cells[grandTotalRow - 1, 3].Formula = $"=SUM(D2:D{grandTotalRow - 1})";
        sheet.Cells[grandTotalRow - 1, 2].PutValue("Grand Total:");

        // Use WorkbookDesigner to process smart markers
        WorkbookDesigner designer = new WorkbookDesigner(workbook);
        designer.SetDataSource(dt);
        designer.Process();

        // Save the result
        workbook.Save("SmartMarkerWithFormula.xlsx");
    }
}
