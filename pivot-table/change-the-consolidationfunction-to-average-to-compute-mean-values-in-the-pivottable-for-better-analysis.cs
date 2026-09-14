// Title: Configure a PivotTable data field to use the Average consolidation function with Aspose.Cells for .NET (C#)
// AI Prompts: Create a new workbook, add sample Category and Amount rows, insert a pivot table, and set the Amount data field's ConsolidationFunction to Average using Aspose.Cells. | Modify an existing Aspose.Cells pivot table so that its data field aggregates values with the Average function instead of the default Sum. | Generate an Excel file where the pivot table calculates the mean Amount per Category by applying the Average consolidation function to the data field.
// Common Searches: Aspose.Cells C# set pivot table data field aggregation to average | change consolidation function to average in Aspose.Cells pivot table | average calculation for pivot table data field using Aspose.Cells .NET
// Tags: Aspose.Cells pivot table average consolidation | C# set pivot data field function | Excel pivot mean aggregation Aspose | Aspose.Cells calculate average in pivot | pivot table data field consolidation function .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // The example creates a workbook, populates it with Category and Amount data, adds a pivot table, places Category in the row area, adds Amount as a data field, sets its ConsolidationFunction to Average, refreshes and calculates the pivot, and saves the file as PivotTable_AverageFunction.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data for the pivot table
                cells["A1"].PutValue("Category");
                cells["B1"].PutValue("Amount");
                cells["A2"].PutValue("A");
                cells["B2"].PutValue(100);
                cells["A3"].PutValue("B");
                cells["B3"].PutValue(200);
                cells["A4"].PutValue("A");
                cells["B4"].PutValue(150);
                cells["A5"].PutValue("B");
                cells["B5"].PutValue(250);

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add fields to the pivot table
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");

                // Add the data field and set its consolidation function to Average
                int dataFieldPos = pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");
                PivotField dataField = pivotTable.DataFields[dataFieldPos];
                dataField.Function = ConsolidationFunction.Average;

                // Refresh and calculate the pivot table to apply the changes
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Ensure the output directory exists
                string outputPath = "PivotTable_AverageFunction.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the pivot table:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
