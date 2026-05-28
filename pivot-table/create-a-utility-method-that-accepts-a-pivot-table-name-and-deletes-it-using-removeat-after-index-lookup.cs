using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace PivotTableDemo
{
    public static class PivotTableUtility
    {
        /// <summary>
        /// Deletes a pivot table by its name from the given workbook.
        /// </summary>
        public static void DeletePivotTableByName(Workbook workbook, string pivotTableName)
        {
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                PivotTableCollection pivots = sheet.PivotTables;
                for (int i = 0; i < pivots.Count; i++)
                {
                    if (pivots[i].Name.Equals(pivotTableName, StringComparison.OrdinalIgnoreCase))
                    {
                        pivots.RemoveAt(i);
                        return;
                    }
                }
            }

            throw new ArgumentException($"Pivot table '{pivotTableName}' not found in the workbook.");
        }
    }

    public class Example
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and add sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(80);

                // Add two pivot tables
                int idx1 = sheet.PivotTables.Add("A1:B3", "D1", "PivotTable1");
                int idx2 = sheet.PivotTables.Add("A1:B3", "D10", "PivotTable2");

                // Configure the first pivot table
                PivotTable pt1 = sheet.PivotTables[idx1];
                pt1.AddFieldToArea(PivotFieldType.Row, "Product");
                pt1.AddFieldToArea(PivotFieldType.Data, "Sales");
                pt1.RefreshData();
                pt1.CalculateData();

                // Delete the second pivot table
                PivotTableUtility.DeletePivotTableByName(workbook, "PivotTable2");

                // Verify removal
                Console.WriteLine("Remaining pivot tables count: " + sheet.PivotTables.Count);

                // Save the workbook
                string outputPath = "Output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }

    public class Program
    {
        public static void Main()
        {
            Example.Run();
        }
    }
}