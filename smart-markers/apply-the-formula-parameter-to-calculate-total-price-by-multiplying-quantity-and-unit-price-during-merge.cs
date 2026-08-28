// Title: Apply a structured reference formula to compute total price after merging a DataTable with smart markers using Aspose.Cells for .NET
// AI Prompts: Write C# code that merges a DataTable into an Excel worksheet with smart markers and then adds a Total column that multiplies Quantity and UnitPrice using a structured reference formula. | Show how to assign a formula to a ListObject column after WorkbookDesigner.Process() so each row calculates its total price in Aspose.Cells. | Demonstrate recalculating all formulas and saving the workbook after inserting a calculated Total column into a smart‑marker generated table.
// Common Searches: asp.net calculate total column with smart markers after workbookdesigner merge | how to set ListColumn.Formula for a listobject created by smart markers in Aspose.Cells | add calculated total field to Excel table generated from a DataTable using Aspose.Cells C# | apply structured reference formula to expanded smart‑marker table in .NET | merge datatable into Excel and compute row total with Aspose.Cells workbookdesigner
// Tags: smart markers listobject column formula | structured reference total calculation Aspose.Cells | WorkbookDesigner merge datatable with calculated column | C# Aspose.Cells set ListColumn formula | excel total price calculation using Aspose.Cells

using System;
using System.Data;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example creates a DataTable of products, defines a worksheet template with smart markers, merges the data via WorkbookDesigner, assigns a structured reference formula ([@Quantity]*[@UnitPrice]) to the Total column of the resulting ListObject, recalculates all formulas, and saves the workbook as MergedWithTotal.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // ---------- Create data source ----------
            DataTable dt = new DataTable("Products");
            dt.Columns.Add("Product", typeof(string));
            dt.Columns.Add("Quantity", typeof(int));
            dt.Columns.Add("UnitPrice", typeof(double));

            dt.Rows.Add("Apple", 5, 1.2);
            dt.Rows.Add("Banana", 3, 0.8);
            dt.Rows.Add("Cherry", 10, 0.5);

            // ---------- Create a workbook template ----------
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Smart markers for merging data
            ws.Cells["A1"].PutValue("&=$Product");
            ws.Cells["B1"].PutValue("&=$Quantity");
            ws.Cells["C1"].PutValue("&=$UnitPrice");
            ws.Cells["D1"].PutValue("Total"); // header for total column

            // Define a table that will expand when data is merged
            int tableIdx = ws.ListObjects.Add("A1", "D1", true);
            ListObject table = ws.ListObjects[tableIdx];
            // Header row is shown by default; explicit setting not required

            // ---------- Merge data using WorkbookDesigner ----------
            WorkbookDesigner designer = new WorkbookDesigner(wb);
            designer.SetDataSource(dt);
            designer.Process();

            // ---------- Apply formula to calculate total price ----------
            // The Total column is the fourth column (index 3) in the table
            ListColumn totalColumn = table.ListColumns[3];
            // Structured reference formula: multiply Quantity and UnitPrice for each row
            totalColumn.Formula = "=[@Quantity]*[@UnitPrice]";

            // ---------- Calculate all formulas ----------
            wb.CalculateFormula();

            // ---------- Save the workbook ----------
            string outputPath = "MergedWithTotal.xlsx";
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
