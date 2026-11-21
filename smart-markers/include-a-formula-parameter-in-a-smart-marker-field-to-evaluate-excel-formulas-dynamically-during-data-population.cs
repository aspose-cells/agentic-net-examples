// Title: Insert a dynamic SUM formula using a Formula smart marker during Aspose.Cells smart marker processing in C#
// Common Searches: Aspose.Cells C# smart marker formula to calculate total after data population | use ${Formula} smart marker for SUM over a dynamic range in an Excel workbook | populate rows with smart markers and automatically evaluate SUM formula in .NET | set RowCount placeholder for dynamic range in Aspose.Cells formula smart marker | process smart markers and compute totals without manual Excel formulas in C#
// Tags: Aspose.Cells formula smart marker | dynamic SUM range using RowCount placeholder | C# smart marker data binding with formulas | Excel total calculation via smart markers | WorkbookDesigner formula evaluation

using Aspose.Cells;
using System.Collections.Generic;

// The example creates a workbook, adds headers, places smart markers for Name, Quantity, and Price, and inserts a Formula smart marker (${Formula=SUM(C2:C${RowCount})}) that computes the total after the data rows are populated. It binds a list of Product objects and a RowCount value to the designer, processes the smart markers to fill the data and evaluate the formula, and saves the result as SmartMarkerWithFormula.xlsx.
public class Product
{
    public string Name { get; set; }
    public int Quantity { get; set; }
    public double Price { get; set; }
}

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Set up column headers
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Quantity");
        sheet.Cells["C1"].PutValue("Price");

        // Insert smart markers for data rows (starting at row 2)
        sheet.Cells["A2"].PutValue("${Name}");
        sheet.Cells["B2"].PutValue("${Quantity}");
        sheet.Cells["C2"].PutValue("${Price}");

        // Insert a label for total
        sheet.Cells["B5"].PutValue("Total");

        // Insert a formula smart marker.
        // The formula will be evaluated after data population.
        // ${RowCount} will be replaced with the index of the last data row.
        sheet.Cells["C5"].PutValue("${Formula=SUM(C2:C${RowCount})}");

        // Prepare sample data
        List<Product> products = new List<Product>()
        {
            new Product(){ Name = "Apple",  Quantity = 10, Price = 1.20 },
            new Product(){ Name = "Banana", Quantity = 5,  Price = 0.80 },
            new Product(){ Name = "Orange", Quantity = 8,  Price = 1.50 }
        };

        // Calculate the last row index that will contain data (header row + data rows)
        int lastDataRow = products.Count + 1; // +1 because data starts at row 2

        // Set up the smart marker processor
        WorkbookDesigner designer = new WorkbookDesigner(workbook);
        designer.SetDataSource("Name", products);
        designer.SetDataSource("Quantity", products);
        designer.SetDataSource("Price", products);
        designer.SetDataSource("RowCount", lastDataRow);
        designer.Process(); // Populate smart markers and evaluate the formula

        // Save the result
        workbook.Save("SmartMarkerWithFormula.xlsx");
    }
}
