using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // ------------------------------------------------------------
        // 1. Populate data for an advanced filter demonstration
        // ------------------------------------------------------------
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Price");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue(1.5);
        sheet.Cells["A3"].PutValue("Banana");
        sheet.Cells["B3"].PutValue(0.8);
        sheet.Cells["A4"].PutValue("Orange");
        sheet.Cells["B4"].PutValue(1.2);
        sheet.Cells["A5"].PutValue("Grape");
        sheet.Cells["B5"].PutValue(2.0);

        // Criteria range: filter only Apple and Orange
        sheet.Cells["D1"].PutValue("Product");
        sheet.Cells["D2"].PutValue("Apple");
        sheet.Cells["D3"].PutValue("Orange");

        // Apply the advanced filter and copy results to E1
        sheet.AdvancedFilter(true, "A1:B5", "D1:D3", "E1", false);

        // Retrieve and display the filter settings
        AdvancedFilter filter = sheet.GetAdvancedFilter();
        Console.WriteLine("List Range: " + filter.ListRange);
        Console.WriteLine("Criteria Range: " + filter.CriteriaRange);
        Console.WriteLine("Copy To Range: " + filter.CopyToRange);

        // ------------------------------------------------------------
        // 2. Add AboveAverage conditional formatting
        // ------------------------------------------------------------
        // Add a numeric column (Score) to demonstrate conditional formatting
        sheet.Cells["C1"].PutValue("Score");
        for (int i = 2; i <= 11; i++)
        {
            sheet.Cells[i, 2].PutValue(i * 10); // Values 20,30,...,110
        }

        // Create a conditional formatting collection
        int cfIndex = sheet.ConditionalFormattings.Add();
        FormatConditionCollection fcs = sheet.ConditionalFormattings[cfIndex];

        // Define the area to which the formatting applies (C2:C11)
        CellArea area = new CellArea
        {
            StartRow = 1,
            EndRow = 10,
            StartColumn = 2,
            EndColumn = 2
        };
        fcs.AddArea(area);

        // Add an AboveAverage condition
        int conditionIdx = fcs.AddCondition(FormatConditionType.AboveAverage);
        FormatCondition fc = fcs[conditionIdx];
        fc.Style.BackgroundColor = Color.Yellow;
        fc.AboveAverage.IsAboveAverage = true;   // Highlight values above average
        fc.AboveAverage.IsEqualAverage = false; // Do not highlight values equal to average

        // ------------------------------------------------------------
        // 3. Demonstrate calculation precision handling
        // ------------------------------------------------------------
        sheet.Cells["F1"].PutValue("Formula");
        sheet.Cells["F2"].Formula = "=-0.45+0.43+0.02"; // Expected result: 0

        // Set calculation options to use the Round precision strategy
        CalculationOptions calcOptions = new CalculationOptions
        {
            PrecisionStrategy = CalculationPrecisionStrategy.Round,
            IgnoreError = true,
            Recursive = true
        };
        workbook.CalculateFormula(calcOptions);
        Console.WriteLine("Rounded formula result: " + sheet.Cells["F2"].DoubleValue);

        // ------------------------------------------------------------
        // Save the workbook
        // ------------------------------------------------------------
        workbook.Save("AdvancedTopicsDemo.xlsx");
    }
}