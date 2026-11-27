// Title: How to verify that Excel chart data labels update after modifying linked cells using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# check chart data label values after changing linked cell range | validate Excel column chart custom data labels programmatically with Aspose.Cells | C# unit test for chart data labels linked to cells using Aspose.Cells | ensure chart data labels reflect updated cell contents in Aspose.Cells .NET | how to programmatically confirm linked data label text changes in an Excel chart with Aspose.Cells
// Tags: chart data labels linked source Aspose.Cells | validate chart label values C# Aspose.Cells | update linked cells Excel chart Aspose.Cells | programmatic data label verification Aspose.Cells | column chart custom labels Aspose.Cells .NET

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook with a column chart whose data labels are linked to cells C2:C4, validates the initial label texts, updates those cells, forces recalculation, and then confirms that the chart data labels reflect the modified values using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ---------- Step 1: Populate source data ----------
            // Category labels
            cells["A1"].PutValue("Category");
            cells["A2"].PutValue("A");
            cells["A3"].PutValue("B");
            cells["A4"].PutValue("C");

            // Numeric values for the chart
            cells["B1"].PutValue("Value");
            cells["B2"].PutValue(10);
            cells["B3"].PutValue(20);
            cells["B4"].PutValue(30);

            // Linked source for data labels (custom text)
            cells["C2"].PutValue("Ten units");
            cells["C3"].PutValue("Twenty units");
            cells["C4"].PutValue("Thirty units");

            // ---------- Step 2: Create a column chart ----------
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Configure data labels to show linked cell range
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = false;          // hide default value
            series.DataLabels.ShowCellRange = true;       // enable linked source
            series.DataLabels.LinkedSource = "C2:C4";     // link to custom text cells
            series.DataLabels.Font.Color = Color.Blue;

            // Save initial workbook (optional visual check)
            workbook.Save("Chart_With_DataLabels_Initial.xlsx", SaveFormat.Xlsx);

            // ---------- Step 3: Validate initial data label texts ----------
            Console.WriteLine("=== Initial Data Labels ===");
            ValidateDataLabels(cells, "C2:C4", new[] { "Ten units", "Twenty units", "Thirty units" });

            // ---------- Step 4: Modify the linked source cells ----------
            cells["C2"].PutValue("Ten updated");
            cells["C3"].PutValue("Twenty updated");
            cells["C4"].PutValue("Thirty updated");

            // Force recalculation (not strictly required for plain text, but safe)
            workbook.CalculateFormula();

            // Save after modification (optional visual check)
            workbook.Save("Chart_With_DataLabels_Modified.xlsx", SaveFormat.Xlsx);

            // ---------- Step 5: Validate data label texts after modification ----------
            Console.WriteLine("\n=== Data Labels After Modification ===");
            ValidateDataLabels(cells, "C2:C4", new[] { "Ten updated", "Twenty updated", "Thirty updated" });

            Console.WriteLine("\nValidation completed successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    /// <param name="cells">The worksheet cells collection.</param>
    /// <param name="range">A range in A1 style (e.g., "C2:C4").</param>
    /// <param name="expectedLabels">Array of expected label strings in order.</param>
    private static void ValidateDataLabels(Cells cells, string range, string[] expectedLabels)
    {
        if (cells == null) throw new ArgumentNullException(nameof(cells));
        if (expectedLabels == null) throw new ArgumentNullException(nameof(expectedLabels));
        if (string.IsNullOrWhiteSpace(range)) throw new ArgumentException("Range cannot be null or empty.", nameof(range));

        // Parse the range (assumes a single column range like C2:C4)
        string[] parts = range.Split(':');
        if (parts.Length != 2) throw new ArgumentException("Range must be in A1:A2 format.", nameof(range));

        CellArea area = CellArea.CreateCellArea(parts[0], parts[1]);

        int rowCount = area.EndRow - area.StartRow + 1;
        if (rowCount != expectedLabels.Length)
            throw new InvalidOperationException($"Cell count ({rowCount}) does not match expected label count ({expectedLabels.Length}).");

        for (int i = 0; i < rowCount; i++)
        {
            int row = area.StartRow + i;
            int col = area.StartColumn;
            string actual = cells[row, col].StringValue ?? string.Empty;
            string expected = expectedLabels[i];

            Console.WriteLine($"Row {row + 1}: Expected = \"{expected}\", Actual = \"{actual}\"");

            if (!string.Equals(actual, expected, StringComparison.Ordinal))
                throw new InvalidOperationException($"Data label mismatch at row {row + 1}. Expected \"{expected}\", but got \"{actual}\".");
        }
    }
}
