// Title: Hide Zero‑Value Rows in an Aspose.Cells Pivot Table (C#)
// Description: Learn how to create a workbook, build a pivot table on Category and Amount fields, calculate totals, and programmatically hide any row whose aggregated value is zero using Aspose.Cells for .NET. The example refreshes, recalculates, and saves the cleaned report.
// Keywords: Aspose.Cells hide zero rows | C# pivot table hide empty rows | Aspose.Cells filter pivot items | remove zero‑total categories Aspose.Cells | Aspose.Cells PivotItem.IsHidden | Excel pivot table zero values C#
// Common Searches: how to hide zero‑value rows in Aspose.Cells pivot table | C# hide empty categories in Excel pivot using Aspose | Aspose.Cells hide pivot items with zero sum | programmatically filter pivot table rows Aspose.Cells | remove zero total rows from pivot table C#
// Developer Intent: Programmatically hide pivot‑table rows whose summed data field equals zero.
// Use Cases: Generate a clean financial or inventory report by excluding categories with no activity. | Automate Excel exports where zero‑total rows clutter the view, improving readability for end users. | Integrate into batch processing pipelines that produce pivot‑based dashboards without empty categories.
// AI Prompts: Write C# code with Aspose.Cells that creates a pivot table and hides rows where the total amount is zero. | Show how to iterate over PivotItems, compare each item's aggregate value to a dictionary, and set IsHidden for zero‑sum entries. | Explain the sequence: refresh pivot data, calculate, hide items, recalculate, and save the workbook with hidden rows.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Learn how to create a workbook, build a pivot table on Category and Amount fields, calculate totals, and programmatically hide any row whose aggregated value is zero using Aspose.Cells for .NET. The example refreshes, recalculates, and saves the cleaned report.
    public class HideZeroValueRowsInPivot
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];

            // Populate sample data with some zero totals
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Item");
            dataSheet.Cells["C1"].PutValue("Amount");

            dataSheet.Cells["A2"].PutValue("Fruit");
            dataSheet.Cells["B2"].PutValue("Apple");
            dataSheet.Cells["C2"].PutValue(100);

            dataSheet.Cells["A3"].PutValue("Fruit");
            dataSheet.Cells["B3"].PutValue("Banana");
            dataSheet.Cells["C3"].PutValue(0);   // Zero value row

            dataSheet.Cells["A4"].PutValue("Vegetable");
            dataSheet.Cells["B4"].PutValue("Carrot");
            dataSheet.Cells["C4"].PutValue(50);

            dataSheet.Cells["A5"].PutValue("Vegetable");
            dataSheet.Cells["B5"].PutValue("Potato");
            dataSheet.Cells["C5"].PutValue(0);   // Zero value row

            // Compute sums per category to identify zero‑total categories
            var categorySums = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            for (int row = 2; row <= 5; row++)
            {
                string category = dataSheet.Cells[row, 0].StringValue;
                double amount = dataSheet.Cells[row, 2].DoubleValue;
                if (!categorySums.ContainsKey(category))
                    categorySums[category] = 0;
                categorySums[category] += amount;
            }

            // Add a new worksheet for the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

            // Create the pivot table
            int pivotIndex = pivotSheet.PivotTables.Add(
                "=Sheet1!A1:C5",   // source data range
                "A1",               // destination upper‑left cell
                "PivotTable1");     // pivot table name

            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Configure the pivot table: Category as row, Amount as data
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Refresh data and calculate the pivot table so that values are available
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Hide rows whose aggregated value is zero
            PivotField rowField = pivotTable.RowFields[0];
            foreach (PivotItem item in rowField.PivotItems)
            {
                if (categorySums.TryGetValue(item.Name, out double sum) && sum == 0)
                {
                    item.IsHidden = true;
                }
            }

            // Re‑calculate after hiding items
            pivotTable.CalculateData();

            // Save the workbook
            string outputPath = "HideZeroValueRowsPivot.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
