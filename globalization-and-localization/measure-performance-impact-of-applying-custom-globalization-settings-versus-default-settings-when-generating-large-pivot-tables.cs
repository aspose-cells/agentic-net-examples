using System;
using System.Diagnostics;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Settings;

class PivotGlobalizationPerformance
{
    // Size of the dataset for the performance test
    const int RowCount = 100_000;

    static void Main()
    {
        // Measure performance with default globalization settings
        double defaultTime = MeasurePivotCreation(useCustomSettings: false);
        Console.WriteLine($"Default globalization settings: {defaultTime:F2} ms");

        // Measure performance with custom globalization settings
        double customTime = MeasurePivotCreation(useCustomSettings: true);
        Console.WriteLine($"Custom globalization settings: {customTime:F2} ms");
    }

    // Creates a workbook, populates data, creates a pivot table and measures the elapsed time.
    static double MeasurePivotCreation(bool useCustomSettings)
    {
        // Start timing
        Stopwatch sw = Stopwatch.StartNew();

        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet dataSheet = workbook.Worksheets[0];

        // Populate a large dataset
        PopulateSampleData(dataSheet, RowCount);

        // Apply custom globalization settings if requested
        if (useCustomSettings)
        {
            // Create an instance of custom pivot globalization settings
            CustomPivotGlobalizationSettings customSettings = new CustomPivotGlobalizationSettings();

            // Assign the custom settings to the workbook
            workbook.Settings.GlobalizationSettings = new GlobalizationSettings();
            workbook.Settings.GlobalizationSettings.PivotSettings = customSettings;
        }

        // Add a new worksheet for the pivot table
        Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

        // Define the data range for the pivot table (A1:C{RowCount})
        string dataRange = $"A1:C{RowCount}";
        // Add the pivot table to the pivot sheet
        int pivotIndex = pivotSheet.PivotTables.Add(dataRange, "A1", "LargePivot");
        PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

        // Configure the pivot fields
        // Row field: Category (column 0)
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);
        // Column field: SubCategory (column 1)
        pivotTable.AddFieldToArea(PivotFieldType.Column, 1);
        // Data field: Value (column 2)
        pivotTable.AddFieldToArea(PivotFieldType.Data, 2);

        // Refresh and calculate the pivot table to apply all settings
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Stop timing
        sw.Stop();

        // Optionally save the workbook (not required for timing)
        // workbook.Save(useCustomSettings ? "CustomPivot.xlsx" : "DefaultPivot.xlsx");

        return sw.Elapsed.TotalMilliseconds;
    }

    // Populates the worksheet with sample data:
    // Column A: Category (e.g., "Category0" to "Category9")
    // Column B: SubCategory (e.g., "Sub0" to "Sub9")
    // Column C: Random numeric value
    static void PopulateSampleData(Worksheet sheet, int rows)
    {
        Random rnd = new Random(0);
        // Header row
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("SubCategory");
        sheet.Cells["C1"].PutValue("Value");

        for (int i = 2; i <= rows + 1; i++)
        {
            // Cycle through 10 categories and subcategories to create realistic grouping
            string category = "Category" + (i % 10);
            string subCategory = "Sub" + (i % 10);
            double value = rnd.NextDouble() * 1000;

            sheet.Cells[$"A{i}"].PutValue(category);
            sheet.Cells[$"B{i}"].PutValue(subCategory);
            sheet.Cells[$"C{i}"].PutValue(value);
        }
    }

    // Custom pivot globalization settings overriding a few text methods
    class CustomPivotGlobalizationSettings : PivotGlobalizationSettings
    {
        public override string GetTextOfTotal()
        {
            return "Custom Total";
        }

        public override string GetTextOfGrandTotal()
        {
            return "Custom Grand Total";
        }

        public override string GetTextOfAll()
        {
            return "All Items (Custom)";
        }

        public override string GetTextOfRowLabels()
        {
            return "Rows (Custom)";
        }

        public override string GetTextOfColumnLabels()
        {
            return "Columns (Custom)";
        }
    }
}