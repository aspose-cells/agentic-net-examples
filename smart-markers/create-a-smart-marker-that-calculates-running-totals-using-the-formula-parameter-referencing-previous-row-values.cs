// Title: Generate a running total in an Excel file using Aspose.Cells C# smart markers with previous‑row formula
// Common Searches: Aspose.Cells C# smart marker formula for cumulative sum across rows | How to reference the previous row in an Aspose.Cells smart marker expression | Create a running total column with {row-1} placeholder in Aspose.Cells | Using WorkbookDesigner to calculate cumulative totals in Excel via smart markers | Smart marker example for dynamic running total in .NET Excel export
// Tags: Aspose.Cells smart marker cumulative sum | previous row reference in smart marker formula | WorkbookDesigner running total calculation | C# Excel export dynamic formula | smart marker {row-1} usage

using System;
using System.Data;
using Aspose.Cells;

// This example shows how to create an Excel workbook with Aspose.Cells for .NET, insert smart markers that compute a running total by referencing the previous row (using {=B{row-1}+A{row}}), bind a DataTable as the data source, process the markers with WorkbookDesigner, recalculate formulas, and save the file as RunningTotal.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Set column headers
            sheet.Cells["A1"].PutValue("Value");
            sheet.Cells["B1"].PutValue("Running Total");

            // Insert smart markers starting from row 2
            // {value} will be replaced by the data source field "value"
            // {=B{row-1}+A{row}} calculates the running total:
            // previous total (B of previous row) + current value (A of current row)
            sheet.Cells["A2"].PutValue("{value}");
            sheet.Cells["B2"].PutValue("{=B{row-1}+A{row}}");

            // Prepare the data source (a DataTable with a single column "value")
            DataTable data = new DataTable();
            data.Columns.Add("value", typeof(double));
            data.Rows.Add(10);
            data.Rows.Add(20);
            data.Rows.Add(15);
            data.Rows.Add(5);

            // Process the smart markers using WorkbookDesigner
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(data);
            designer.Process();

            // Recalculate formulas to ensure the running totals are computed
            workbook.CalculateFormula();

            // Save the resulting workbook
            string outputPath = "RunningTotal.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
