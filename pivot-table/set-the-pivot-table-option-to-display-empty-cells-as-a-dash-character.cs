using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Settings;

namespace AsposeCellsExamples
{
    public class PivotTableEmptyCellDashDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet dataSheet = workbook.Worksheets[0];

                // Populate sample data (including some empty cells)
                dataSheet.Cells["A1"].PutValue("Category");
                dataSheet.Cells["B1"].PutValue("Value");
                dataSheet.Cells["A2"].PutValue("A");
                dataSheet.Cells["B2"].PutValue(10);
                dataSheet.Cells["A3"].PutValue("B");
                dataSheet.Cells["B3"].PutValue(20);
                dataSheet.Cells["A4"].PutValue("");   // Empty category cell
                dataSheet.Cells["B4"].PutValue(30);

                // Add a pivot table based on the data range
                int pivotIndex = dataSheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
                PivotTable pivotTable = dataSheet.PivotTables[pivotIndex];

                // Configure the pivot table fields
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category
                pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Value

                // Create globalization settings to replace empty data label with a dash ("-")
                SettablePivotGlobalizationSettings pivotSettings = new SettablePivotGlobalizationSettings();
                pivotSettings.SetTextOfEmptyData("-"); // Set dash for empty cells

                // Apply the settings to the workbook
                SettableGlobalizationSettings globalizationSettings = new SettableGlobalizationSettings
                {
                    PivotSettings = pivotSettings
                };
                workbook.Settings.GlobalizationSettings = globalizationSettings;

                // Refresh and calculate the pivot table to apply the changes
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook
                string outputPath = "PivotTableEmptyCellDashDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            PivotTableEmptyCellDashDemo.Run();
        }
    }
}