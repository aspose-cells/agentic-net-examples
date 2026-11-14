// Title: Apply the {row} parameter in Aspose.Cells smart markers to generate per‑row Excel formulas in C#
// Common Searches: how to generate a total column with per‑row formula using Aspose.Cells smart markers in C# | using {row} placeholder in Aspose.Cells smart marker to calculate Qty multiplied by Price | Aspose.Cells WorkbookDesigner apply row‑specific formula for each data row | C# example of smart marker formula with {row} parameter for Excel export | dynamic Excel formula insertion with Aspose.Cells smart markers and DataTable
// Tags: Aspose.Cells smart marker row parameter | C# WorkbookDesigner per‑row formula | Excel export dynamic formula Aspose.Cells | smart marker formula with DataTable source | generate total column using {row} placeholder

using System;
using System.Data;
using Aspose.Cells;

// The example creates a workbook template, places smart markers for Item, Qty, and Price, and uses the ${=Qty{row}*Price{row}} marker to compute a Total column. A DataTable supplies the data, which is bound to WorkbookDesigner, processed to replace the markers, and the resulting Excel file is saved.
class SmartMarkerFormulaExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Template";

            // Set up the header row
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Qty");
            sheet.Cells["C1"].PutValue("Price");
            sheet.Cells["D1"].PutValue("Total");

            // Insert smart markers for data rows starting from row 2
            // ${Item}, ${Qty}, ${Price} will be replaced with data values
            // The formula marker uses the {row} parameter so that each row gets its own formula
            sheet.Cells["A2"].PutValue("${Item}");
            sheet.Cells["B2"].PutValue("${Qty}");
            sheet.Cells["C2"].PutValue("${Price}");
            // Formula: Qty * Price for the current row
            sheet.Cells["D2"].PutValue("${=Qty{row}*Price{row}}");

            // Prepare the data source
            DataTable dt = new DataTable("Data");
            dt.Columns.Add("Item", typeof(string));
            dt.Columns.Add("Qty", typeof(int));
            dt.Columns.Add("Price", typeof(double));

            dt.Rows.Add("Apple", 10, 0.5);
            dt.Rows.Add("Banana", 5, 0.2);
            dt.Rows.Add("Cherry", 20, 0.75);

            // Use WorkbookDesigner to process smart markers
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource("Data", dt); // associate the DataTable with the name "Data"
            designer.Process(); // process smart markers (including formulas)

            // Save the result (lifecycle rule: save)
            string outputPath = "SmartMarkerFormulaResult.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
