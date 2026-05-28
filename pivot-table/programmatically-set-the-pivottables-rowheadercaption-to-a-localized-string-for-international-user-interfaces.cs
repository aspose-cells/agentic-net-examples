using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class PivotTableRowHeaderLocalization
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet dataSheet = workbook.Worksheets[0];
                dataSheet.Name = "Data";

                // Populate sample data
                var cells = dataSheet.Cells;
                cells["A1"].PutValue("Country");
                cells["B1"].PutValue("Sales");
                cells["A2"].PutValue("USA");
                cells["B2"].PutValue(1500);
                cells["A3"].PutValue("Germany");
                cells["B3"].PutValue(1200);
                cells["A4"].PutValue("Japan");
                cells["B4"].PutValue(1800);

                // Add a worksheet to host the pivot table
                Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

                // Create the pivot table
                int pivotIndex = pivotSheet.PivotTables.Add("=Data!A1:B4", "A3", "SalesPivot");
                PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

                // Configure pivot fields
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Country");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Localized row header caption
                string localizedRowHeader = "Land"; // German for "Country"
                pivotTable.RowHeaderCaption = localizedRowHeader;

                // NOTE: Globalization settings API may vary between Aspose.Cells versions.
                // If needed, adjust the following lines to the appropriate method for your version.
                // Example for older versions:
                // SettablePivotGlobalizationSettings globalizationSettings = new SettablePivotGlobalizationSettings();
                // globalizationSettings.SetTextOfRowLabels("Zeilenbeschriftungen"); // German for "Row Labels"
                // workbook.Settings.SetGlobalizationSettings(globalizationSettings);

                // Save the workbook
                workbook.Save("PivotTableRowHeaderLocalized.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point
    class Program
    {
        static void Main(string[] args)
        {
            PivotTableRowHeaderLocalization.Run();
        }
    }
}