// Title: C# Example – Embed a Formula in Aspose.Cells Smart Markers for Dynamic Totals
// Description: This Aspose.Cells .NET sample creates a workbook, adds smart markers for product data, inserts a formula (Price × Quantity) into the Total column, processes the markers with WorkbookDesigner, recalculates the formula, and saves the result as an Excel file.
// Keywords: Aspose.Cells C# smart markers formula | embed Excel formula Aspose.Cells | dynamic calculation smart markers .NET | WorkbookDesigner Process CalculateFormula | C# Excel total column example | Aspose.Cells sample GitHub
// Common Searches: how to add a formula that uses smart marker values in Aspose.Cells C# | calculate totals after processing smart markers .NET | embed Excel formulas with Aspose.Cells smart markers | Aspose.Cells dynamic calculations example | C# smart marker formula without leading equals
// Developer Intent: Insert a cell formula that references smart‑marker fields so totals are computed automatically after the data is populated.
// Use Cases: Place a formula (e.g., B2*C2) in a cell before calling WorkbookDesigner.Process so it automatically uses the values filled by smart markers. | Bind a collection of objects as the smart‑marker data source, run Process, then invoke Workbook.CalculateFormula to evaluate all embedded formulas. | Generate an Excel report where the Total column shows calculated results for each product row without manual post‑processing.
// AI Prompts: Generate C# code that adds a formula referencing smart‑marker cells and evaluates it after processing with Aspose.Cells. | Explain step‑by‑step how to embed and recalculate an Excel formula when using smart markers in Aspose.Cells for .NET. | Provide a complete Aspose.Cells example that creates a workbook, inserts smart markers, embeds a dependent formula, processes the data source, and saves the calculated file.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace SmartMarkerFormulaDemo
{
    // Simple data class for the smart marker data source
    // This Aspose.Cells .NET sample creates a workbook, adds smart markers for product data, inserts a formula (Price × Quantity) into the Total column, processes the markers with WorkbookDesigner, recalculates the formula, and saves the result as an Excel file.
    public class Product
    {
        public string Name { get; set; } = string.Empty;
        public double Price { get; set; }
        public int Qty { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // ---------- Create a new workbook ----------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // ---------- Define header row ----------
                cells["A1"].PutValue("Product");
                cells["B1"].PutValue("Price");
                cells["C1"].PutValue("Quantity");
                cells["D1"].PutValue("Total");

                // ---------- Insert smart markers ----------
                // These markers will be replaced by the data source during processing
                cells["A2"].PutValue("&=Products.Name");
                cells["B2"].PutValue("&=Products.Price");
                cells["C2"].PutValue("&=Products.Qty");

                // Embed a formula that references the cells filled by smart markers.
                // The formula calculates Total = Price * Quantity.
                cells["D2"].Formula = "B2*C2"; // Set formula without leading '='

                // ---------- Prepare data source ----------
                List<Product> productList = new List<Product>
                {
                    new Product { Name = "Apple",  Price = 1.20, Qty = 10 },
                    new Product { Name = "Banana", Price = 0.80, Qty = 15 },
                    new Product { Name = "Cherry", Price = 2.50, Qty = 5 }
                };

                // ---------- Process smart markers ----------
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource("Products", productList);
                designer.Process(); // Populate smart markers with data

                // ---------- Calculate formulas ----------
                // After the data is populated, calculate the embedded formula.
                workbook.CalculateFormula();

                // ---------- Save the result ----------
                workbook.Save("SmartMarkerWithFormula.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
