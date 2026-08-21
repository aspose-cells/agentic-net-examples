// Title: Create a new worksheet, copy source data, and add a linked PivotTable with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that uses Aspose.Cells to copy a range from a source worksheet to a newly added worksheet and then creates a linked PivotTable on a third worksheet referencing the copied range. | Show how to adapt the example to open an existing workbook, copy its data to another sheet, and define the PivotTable source as a dynamic named range.
// Common Searches: asp.net aspose.cells copy data to new worksheet and create linked pivot table | c# programmatically add worksheet and linked pivot table using Aspose.Cells | how to reference a copied range in a PivotTable with Aspose.Cells .NET | Aspose.Cells example for creating a PivotTable that uses data from another sheet | C# Aspose.Cells create linked pivot table from copied data range
// Tags: Aspose.Cells copy worksheet range C# | Aspose.Cells linked pivot table creation | Aspose.Cells add worksheet and pivot table | Aspose.Cells set pivot source to another sheet | Aspose.Cells refresh pivot data programmatically

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsPivotExample
{
    // The program creates a workbook, adds a source worksheet with sample data, copies that data to a new worksheet, then adds a third worksheet containing a linked PivotTable that references the copied range, configures row and data fields, refreshes and calculates the pivot, and saves the file as LinkedPivotTableDemo.xlsx.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // ---------- Source Worksheet ----------
                Worksheet sourceSheet = workbook.Worksheets[0];
                sourceSheet.Name = "SourceData";

                // Populate sample source data
                sourceSheet.Cells["A1"].PutValue("Category");
                sourceSheet.Cells["B1"].PutValue("Value");
                sourceSheet.Cells["A2"].PutValue("A");
                sourceSheet.Cells["B2"].PutValue(10);
                sourceSheet.Cells["A3"].PutValue("B");
                sourceSheet.Cells["B3"].PutValue(20);
                sourceSheet.Cells["A4"].PutValue("A");
                sourceSheet.Cells["B4"].PutValue(30);

                // ---------- Destination Worksheet (copy of source) ----------
                Worksheet copySheet = workbook.Worksheets.Add("CopyData");

                // Determine the used range in the source sheet
                AsposeRange sourceRange = sourceSheet.Cells.MaxDisplayRange;

                // Copy the source range to the destination sheet starting at A1
                int rowCount = sourceRange.RowCount;
                int colCount = sourceRange.ColumnCount;
                AsposeRange destRange = copySheet.Cells.CreateRange(0, 0, rowCount, colCount);
                sourceRange.Copy(destRange);

                // ---------- PivotTable Worksheet ----------
                Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

                // Build the source data reference string for the copied data
                AsposeRange copiedRange = copySheet.Cells.MaxDisplayRange;
                string sourceData = $"=CopyData!{copiedRange.Address}";

                // Add a linked PivotTable that uses the copied data as its source
                PivotTableCollection pivotTables = pivotSheet.PivotTables;
                int pivotIndex = pivotTables.Add(sourceData, "A1", "LinkedPivot");

                // Configure the PivotTable (optional but typical)
                PivotTable pivotTable = pivotTables[pivotIndex];
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Value");

                // Refresh and calculate the PivotTable data
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook
                workbook.Save("LinkedPivotTableDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
