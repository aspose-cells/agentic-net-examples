// Title: Create a linked PivotTable on a separate sheet after copying data with Aspose.Cells for .NET
// Description: A C# example that builds a workbook, copies a used range to a new worksheet while preserving styles, adds another sheet, and inserts a PivotTable that references the copied data. The PivotTable is configured, refreshed, calculated, and the file is saved as LinkedPivotTableDemo.xlsx.
// Keywords: Aspose.Cells PivotTable C# | copy range to another worksheet Aspose.Cells | linked PivotTable source reference | refresh PivotTable programmatically | save workbook with PivotTable | .NET Excel automation | Excel pivot from copied data
// Common Searches: Aspose.Cells copy worksheet range and create PivotTable | C# linked PivotTable on different sheet Aspose.Cells | How to reference another sheet in Aspose.Cells PivotTable | Programmatic PivotTable refresh Aspose.Cells .NET | Example of copying data and building a PivotTable with Aspose
// Developer Intent: Copy source data to a new sheet and generate a PivotTable that points to the copied range using Aspose.Cells for .NET.
// Use Cases: Produce a reporting sheet that isolates raw data from analysis tables. | Automate Excel workbooks where original data must remain unchanged while pivot reports are generated. | Refresh PivotTable calculations after programmatic data updates.
// AI Prompts: Show C# code with Aspose.Cells to copy a range from one worksheet to another and create a linked PivotTable referencing the copied range. | Explain how to assign row, column, and data fields in an Aspose.Cells PivotTable and refresh its data via code.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsPivotExample
{
    // A C# example that builds a workbook, copies a used range to a new worksheet while preserving styles, adds another sheet, and inserts a PivotTable that references the copied data. The PivotTable is configured, refreshed, calculated, and the file is saved as LinkedPivotTableDemo.xlsx.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook (default contains one worksheet)
                Workbook workbook = new Workbook();

                // -------------------------------------------------
                // 1. Prepare source data in the first worksheet
                // -------------------------------------------------
                Worksheet sourceSheet = workbook.Worksheets[0];
                sourceSheet.Name = "SourceData";

                // Sample data
                sourceSheet.Cells["A1"].PutValue("Category");
                sourceSheet.Cells["B1"].PutValue("Product");
                sourceSheet.Cells["C1"].PutValue("Sales");

                sourceSheet.Cells["A2"].PutValue("Fruit");
                sourceSheet.Cells["B2"].PutValue("Apple");
                sourceSheet.Cells["C2"].PutValue(120);

                sourceSheet.Cells["A3"].PutValue("Fruit");
                sourceSheet.Cells["B3"].PutValue("Orange");
                sourceSheet.Cells["C3"].PutValue(150);

                sourceSheet.Cells["A4"].PutValue("Vegetable");
                sourceSheet.Cells["B4"].PutValue("Carrot");
                sourceSheet.Cells["C4"].PutValue(80);

                // -------------------------------------------------
                // 2. Add a new worksheet that will hold the copied data
                // -------------------------------------------------
                Worksheet dataCopySheet = workbook.Worksheets.Add("DataCopy");

                // Determine the used range of the source sheet
                AsposeRange usedRange = sourceSheet.Cells.MaxDisplayRange;

                // Copy values and styles from source to destination sheet
                for (int r = 0; r < usedRange.RowCount; r++)
                {
                    for (int c = 0; c < usedRange.ColumnCount; c++)
                    {
                        Cell srcCell = sourceSheet.Cells[usedRange.FirstRow + r, usedRange.FirstColumn + c];
                        Cell destCell = dataCopySheet.Cells[usedRange.FirstRow + r, usedRange.FirstColumn + c];
                        destCell.PutValue(srcCell.Value);
                        destCell.SetStyle(srcCell.GetStyle());
                    }
                }

                // -------------------------------------------------
                // 3. Add a worksheet that will contain the PivotTable
                // -------------------------------------------------
                Worksheet pivotSheet = workbook.Worksheets.Add("PivotReport");

                // Build the source data reference string for the PivotTable (e.g., =DataCopy!A1:C4)
                string sourceData = $"=DataCopy!{usedRange.Address}";

                // Add a new PivotTable to the pivot sheet (destination cell A1, name "SalesPivot")
                PivotTableCollection pivotTables = pivotSheet.PivotTables;
                int pivotIndex = pivotTables.Add(sourceData, "A1", "SalesPivot");

                // Configure the PivotTable fields
                PivotTable pivotTable = pivotTables[pivotIndex];
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Column, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Refresh and calculate the PivotTable data
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // -------------------------------------------------
                // 4. Save the workbook
                // -------------------------------------------------
                workbook.Save("LinkedPivotTableDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
