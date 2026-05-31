using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class AddCalculatedFieldConcatenateDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data with two text columns
                cells["A1"].Value = "First";
                cells["B1"].Value = "Second";

                cells["A2"].Value = "Apple";
                cells["B2"].Value = "Red";

                cells["A3"].Value = "Banana";
                cells["B3"].Value = "Yellow";

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B3", "D1", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add source fields to the pivot (optional, just to have a visible pivot)
                pivotTable.AddFieldToArea(PivotFieldType.Row, "First");
                pivotTable.AddFieldToArea(PivotFieldType.Column, "Second");
                // Add a dummy data field so the pivot has a data area (required for calculated field)
                pivotTable.AddFieldToArea(PivotFieldType.Data, "First");

                // Add a calculated field that concatenates the two text fields with a hyphen
                // Formula syntax: =First & "-" & Second
                pivotTable.AddCalculatedField("First-Second", "=First & \"-\" & Second", true);

                // Refresh and calculate the pivot table to apply the new field
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook
                workbook.Save("PivotTable_With_ConcatenatedField.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                // Log any runtime errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            AddCalculatedFieldConcatenateDemo.Run();
        }
    }
}