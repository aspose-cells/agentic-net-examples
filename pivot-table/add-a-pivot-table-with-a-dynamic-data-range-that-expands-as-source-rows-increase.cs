// Title: Add a pivot table with a dynamic data range that expands as source rows increase using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that creates a pivot table whose source range references whole columns so the range grows automatically when new rows are added. | Show how to append additional rows to the source worksheet and then refresh and recalculate the pivot table to reflect the new data.
// Common Searches: Aspose.Cells C# create pivot table with dynamic source range using whole column reference | how to make a pivot table auto‑expand when adding rows in Aspose.Cells .NET | refresh pivot table after adding data rows with Aspose.Cells C# | set pivot table source to A:B column range in Aspose.Cells | dynamic pivot table example Aspose.Cells workbook C#
// Tags: dynamic source range for Aspose.Cells pivot table | refresh pivot table after data addition Aspose.Cells | whole column reference pivot source Aspose.Cells | C# Aspose.Cells create expanding pivot table | pivot table auto‑expand rows Aspose.Cells .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsDynamicPivot
{
    // The program creates a workbook, adds a source worksheet with sample data, defines a dynamic source range using whole columns (A:B), inserts a pivot table on a separate sheet, appends extra rows to the source, refreshes and recalculates the pivot, and saves the file as DynamicPivotTableDemo.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // 1. Prepare source data worksheet (named "SourceData")
            // -------------------------------------------------
            Worksheet sourceSheet = workbook.Worksheets[0];
            sourceSheet.Name = "SourceData";

            // Add header row
            sourceSheet.Cells["A1"].PutValue("Category");
            sourceSheet.Cells["B1"].PutValue("Value");

            // Add initial sample rows
            sourceSheet.Cells["A2"].PutValue("A");
            sourceSheet.Cells["B2"].PutValue(10);
            sourceSheet.Cells["A3"].PutValue("B");
            sourceSheet.Cells["B3"].PutValue(20);
            sourceSheet.Cells["A4"].PutValue("A");
            sourceSheet.Cells["B4"].PutValue(30);

            // -------------------------------------------------
            // 2. Create a worksheet that will hold the pivot table
            // -------------------------------------------------
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

            // -------------------------------------------------
            // 3. Define a dynamic source range.
            //    Using whole columns (A:B) ensures the range expands
            //    automatically when new rows are added to the source sheet.
            // -------------------------------------------------
            string sourceData = "=SourceData!A:B";

            // -------------------------------------------------
            // 4. Add the pivot table using the Add(string, string, string) overload
            // -------------------------------------------------
            PivotTableCollection pivotTables = pivotSheet.PivotTables;
            int pivotIndex = pivotTables.Add(sourceData, "A1", "MyPivotTable");

            // -------------------------------------------------
            // 5. Configure the pivot table (Category as Row, Value as Data)
            // -------------------------------------------------
            PivotTable pivotTable = pivotTables[pivotIndex];
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Value");

            // -------------------------------------------------
            // 6. (Optional) Demonstrate that the pivot updates when source grows
            // -------------------------------------------------
            // Add more rows to the source data after the pivot has been created
            sourceSheet.Cells["A5"].PutValue("C");
            sourceSheet.Cells["B5"].PutValue(40);
            sourceSheet.Cells["A6"].PutValue("B");
            sourceSheet.Cells["B6"].PutValue(25);

            // Refresh the pivot table so it reflects the new rows
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // -------------------------------------------------
            // 7. Save the workbook
            // -------------------------------------------------
            workbook.Save("DynamicPivotTableDemo.xlsx");
        }
    }
}
