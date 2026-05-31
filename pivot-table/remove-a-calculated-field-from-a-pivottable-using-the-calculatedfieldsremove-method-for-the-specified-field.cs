using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class RemoveCalculatedFieldDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                worksheet.Cells["A1"].PutValue("Product");
                worksheet.Cells["B1"].PutValue("Sales");
                worksheet.Cells["C1"].PutValue("Cost");

                worksheet.Cells["A2"].PutValue("Apple");
                worksheet.Cells["B2"].PutValue(1200);
                worksheet.Cells["C2"].PutValue(800);

                worksheet.Cells["A3"].PutValue("Banana");
                worksheet.Cells["B3"].PutValue(1500);
                worksheet.Cells["C3"].PutValue(900);

                worksheet.Cells["A4"].PutValue("Orange");
                worksheet.Cells["B4"].PutValue(1800);
                worksheet.Cells["C4"].PutValue(1100);

                // Add a pivot table based on the data range
                int pivotIndex = worksheet.PivotTables.Add("A1:C4", "E3", "PivotTable1");
                PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

                // Add fields to the pivot table
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Cost");

                // Add a calculated field (Profit = Sales - Cost) and drag it to the data area
                string calcFieldName = "Profit";
                pivotTable.AddCalculatedField(calcFieldName, "Sales-Cost", true);

                // Refresh and calculate to ensure the calculated field appears
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Verify that the calculated field exists in the DataFields collection
                bool hasCalcField = false;
                foreach (PivotField df in pivotTable.DataFields)
                {
                    if (df.Name == calcFieldName && df.IsCalculatedField)
                    {
                        hasCalcField = true;
                        break;
                    }
                }

                Console.WriteLine($"Calculated field '{calcFieldName}' added: {hasCalcField}");

                // Remove the calculated field using RemoveField (field type is Data)
                pivotTable.RemoveField(PivotFieldType.Data, calcFieldName);

                // Recalculate after removal
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Verify removal
                bool stillExists = false;
                foreach (PivotField df in pivotTable.DataFields)
                {
                    if (df.Name == calcFieldName && df.IsCalculatedField)
                    {
                        stillExists = true;
                        break;
                    }
                }

                Console.WriteLine($"Calculated field '{calcFieldName}' present after removal: {stillExists}");

                // Save the workbook to a file
                string outputPath = "PivotTable_RemoveCalculatedField.xlsx";
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error saving workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            RemoveCalculatedFieldDemo.Run();
        }
    }
}