// Title: Consolidate data from multiple worksheets into a PivotTable with Aspose.Cells for .NET
// Description: C# example that creates a workbook, fills identical tables on two sheets, applies a sum subtotal via the ConsolidationFunction property, and builds a consolidated PivotTable on a third sheet using multiple source ranges. The PivotTable shows Category as rows and Value as data, then saves the file as an Excel workbook.
// Keywords: Aspose.Cells PivotTable multiple worksheets | ConsolidationFunction.Sum C# | consolidated pivot table Aspose.Cells | .NET Excel pivot from several sheets | add pivot table with multiple source ranges | subtotal using ConsolidationFunction
// Common Searches: how to create a pivot table from several worksheets using Aspose.Cells | Aspose.Cells ConsolidationFunction sum subtotal example | C# combine data from multiple sheets into one pivot table | Aspose.Cells add pivot table with multiple ranges | consolidated pivot report Aspose.Cells .NET
// Developer Intent: Generate a PivotTable that aggregates data across multiple worksheets and applies a sum subtotal using the ConsolidationFunction property in Aspose.Cells for .NET.
// Use Cases: Summarize regional sales figures from separate worksheets into a single pivot report. | Aggregate inventory counts from multiple location sheets for a company‑wide overview. | Produce a financial consolidation where department totals are summed into one pivot table.
// AI Prompts: Show how to add a column field (e.g., Date) to the consolidated PivotTable. | Provide code to export the consolidated PivotTable to PDF with Aspose.Cells. | Explain how to switch the ConsolidationFunction from Sum to Average for the subtotals.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // C# example that creates a workbook, fills identical tables on two sheets, applies a sum subtotal via the ConsolidationFunction property, and builds a consolidated PivotTable on a third sheet using multiple source ranges. The PivotTable shows Category as rows and Value as data, then saves the file as an Excel workbook.
    public class ConsolidatedPivotTableDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // -------------------------------------------------
                // Prepare sample data in two worksheets (Sheet1 and Sheet2)
                // -------------------------------------------------
                Worksheet sheet1 = workbook.Worksheets[0];
                sheet1.Name = "Sheet1";
                FillSampleData(sheet1);

                Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
                FillSampleData(sheet2);

                // -------------------------------------------------
                // Apply subtotal on each sheet using ConsolidationFunction.Sum
                // -------------------------------------------------
                // Define the data area (A1:C5) for subtotal
                CellArea dataArea = new CellArea
                {
                    StartRow = 0,
                    StartColumn = 0,
                    EndRow = 4,
                    EndColumn = 2
                };
                // Apply subtotal on the "Value" column (index 1) with Sum function
                sheet1.Cells.Subtotal(dataArea, 0, ConsolidationFunction.Sum, new int[] { 1 });
                sheet2.Cells.Subtotal(dataArea, 0, ConsolidationFunction.Sum, new int[] { 1 });

                // -------------------------------------------------
                // Create a new worksheet to host the consolidated PivotTable
                // -------------------------------------------------
                Worksheet pivotSheet = workbook.Worksheets.Add("ConsolidatedPivot");

                // -------------------------------------------------
                // Define multiple consolidation ranges (the two sheets)
                // -------------------------------------------------
                string[] sourceRanges = new string[]
                {
                    "=Sheet1!A1:C5",
                    "=Sheet2!A1:C5"
                };

                // Create empty page fields (not using auto page)
                PivotPageFields pageFields = new PivotPageFields();

                // Add the PivotTable using the overload that accepts multiple ranges
                PivotTableCollection pivotTables = pivotSheet.PivotTables;
                int pivotIndex = pivotTables.Add(sourceRanges, false, pageFields, "A1", "ConsolidatedPivotTable");

                // -------------------------------------------------
                // Configure the PivotTable (Row: Category, Data: Value)
                // -------------------------------------------------
                PivotTable pivotTable = pivotTables[pivotIndex];
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Value");

                // -------------------------------------------------
                // Save the workbook
                // -------------------------------------------------
                workbook.Save("ConsolidatedPivotTableDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during execution: {ex.Message}");
            }
        }

        // Helper method to fill sample data into a worksheet
        private static void FillSampleData(Worksheet sheet)
        {
            Cells cells = sheet.Cells;
            // Header
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Value");
            cells["C1"].PutValue("Date");

            // Sample rows
            for (int i = 0; i < 5; i++)
            {
                cells[$"A{i + 2}"].PutValue($"Item{i + 1}");
                cells[$"B{i + 2}"].PutValue((i + 1) * 100);
                cells[$"C{i + 2}"].PutValue(DateTime.Now.AddDays(i).ToString("yyyy-MM-dd"));
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ConsolidatedPivotTableDemo.Run();
        }
    }
}
