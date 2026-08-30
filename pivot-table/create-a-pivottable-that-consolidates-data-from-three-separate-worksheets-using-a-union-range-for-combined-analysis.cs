// Title: Create a consolidated PivotTable from three worksheets using union ranges with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses Aspose.Cells to define three worksheets, combine their A1:B5 ranges into a union, and generate a PivotTable on a new sheet. | Show how to assign the 'Category' field as a row and the 'Value' field as data in the PivotTable, then refresh and calculate the results programmatically. | Demonstrate saving the workbook as ConsolidatedPivot.xlsx after the PivotTable has been refreshed.
// Common Searches: Aspose.Cells C# create pivot table from multiple worksheets using union range | how to use multiple consolidation ranges overload in Aspose.Cells PivotTable | programmatically add a PivotTable to a separate sheet in Aspose.Cells .NET | refresh and calculate data for a consolidated pivot table with Aspose.Cells | sample code for consolidating sheet data into one pivot table in C#
// Tags: Aspose.Cells multiple consolidation ranges PivotTable | C# union range pivot table creation | add pivot table to new worksheet Aspose.Cells | refresh calculate pivot Aspose.Cells API | save workbook ConsolidatedPivot.xlsx C# | sample data population Aspose.Cells worksheets

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotUnionExample
{
    // The example creates a workbook with three worksheets containing identical sample data, defines a union of the A1:B5 ranges from each sheet, and adds a consolidated PivotTable on a fourth worksheet. It sets 'Category' as the row field and 'Value' as the data field, refreshes and calculates the PivotTable, then saves the file as ConsolidatedPivot.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // -------------------------------------------------
                // Prepare three worksheets with sample data
                // -------------------------------------------------
                Worksheet sheet1 = workbook.Worksheets[0];
                sheet1.Name = "Sheet1";
                FillSampleData(sheet1);

                Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
                FillSampleData(sheet2);

                Worksheet sheet3 = workbook.Worksheets.Add("Sheet3");
                FillSampleData(sheet3);

                // -------------------------------------------------
                // Add a worksheet that will host the consolidated PivotTable
                // -------------------------------------------------
                Worksheet pivotSheet = workbook.Worksheets.Add("PivotSheet");

                // Define the consolidation ranges (union of three sheets)
                // Note: range strings should NOT start with '=' when using the multiple‑range overload.
                string[] sourceRanges = {
                    "Sheet1!A1:B5",
                    "Sheet2!A1:B5",
                    "Sheet3!A1:B5"
                };

                // Create an empty PivotPageFields object (required by the API)
                PivotPageFields pageFields = new PivotPageFields();

                // Add the PivotTable using the multiple consolidation ranges overload
                PivotTableCollection pivotTables = pivotSheet.PivotTables;
                int pivotIndex = pivotTables.Add(sourceRanges, false, pageFields, "A1", "ConsolidatedPivot");
                PivotTable pivot = pivotTables[pivotIndex];

                // Configure the PivotTable fields
                pivot.AddFieldToArea(PivotFieldType.Row, "Category");
                pivot.AddFieldToArea(PivotFieldType.Data, "Value");

                // Refresh and calculate the PivotTable data
                pivot.RefreshData();
                pivot.CalculateData();

                // Save the workbook
                workbook.Save("ConsolidatedPivot.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }

        // Helper method to populate a worksheet with identical sample data
        private static void FillSampleData(Worksheet sheet)
        {
            Cells cells = sheet.Cells;
            // Header
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Value");
            // Sample rows
            cells["A2"].PutValue("A");
            cells["B2"].PutValue(10);
            cells["A3"].PutValue("B");
            cells["B3"].PutValue(20);
            cells["A4"].PutValue("C");
            cells["B4"].PutValue(30);
            cells["A5"].PutValue("A");
            cells["B5"].PutValue(40);
        }
    }
}
