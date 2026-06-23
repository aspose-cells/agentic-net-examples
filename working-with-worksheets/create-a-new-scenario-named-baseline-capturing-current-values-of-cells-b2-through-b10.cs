using Aspose.Cells;
using System;

class CreateBaselineScenario
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // (Optional) Populate cells B2:B10 with sample data
        for (int i = 1; i <= 9; i++) // rows 1..9 correspond to B2..B10
        {
            worksheet.Cells[i, 1].PutValue(i * 10); // example values
        }

        // Add a new scenario named "Baseline"
        ScenarioCollection scenarios = worksheet.Scenarios;
        int scenarioIndex = scenarios.Add("Baseline");
        Scenario baselineScenario = scenarios[scenarioIndex];

        // Capture current values of cells B2 through B10 into the scenario
        for (int row = 1; row <= 9; row++) // B2 (row 1) to B10 (row 9)
        {
            // Retrieve the cell value as a string; use empty string if null
            string cellValue = worksheet.Cells[row, 1].Value?.ToString() ?? string.Empty;

            // Add the input cell to the scenario (row index, column index, value)
            baselineScenario.InputCells.Add(row, 1, cellValue);
        }

        // Save the workbook with the scenario
        workbook.Save("BaselineScenario.xlsx");
    }
}