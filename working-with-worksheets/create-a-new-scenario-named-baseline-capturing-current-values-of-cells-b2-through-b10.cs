// Title: Add a Baseline Scenario for Cells B2‑B10 with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a new workbook, optionally fill range B2:B10 with sample data, add a scenario named "Baseline", capture the current values of those cells as input cells, and save the file as BaselineScenario.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells scenario | C# baseline scenario | capture cell values Aspose.Cells | input cells B2 B10 | save workbook with scenario .NET
// Common Searches: Aspose.Cells create baseline scenario C# | how to add input cells to a scenario in Aspose.Cells | store current values of B2:B10 as a scenario | save workbook with scenarios Aspose.Cells .NET | scenario collection example Aspose.Cells
// Developer Intent: Create a "Baseline" scenario that records the current values of cells B2 through B10.
// Use Cases: Preserve initial data for what‑if analysis by storing it in a baseline scenario. | Enable quick rollback to original values after running calculations or data updates. | Share a workbook with predefined input sets for multiple users or automated tests.
// AI Prompts: Show how to add output cells to the Baseline scenario after calculations. | Explain how to switch between multiple scenarios programmatically with Aspose.Cells. | Provide code to update the Baseline scenario with new values without recreating it.

using System;
using Aspose.Cells;

// Demonstrates how to create a new workbook, optionally fill range B2:B10 with sample data, add a scenario named "Baseline", capture the current values of those cells as input cells, and save the file as BaselineScenario.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // (Optional) Populate B2:B10 with sample data for demonstration
        for (int i = 1; i <= 9; i++) // Row indices 1..9 correspond to B2..B10
        {
            worksheet.Cells[i, 1].PutValue(i * 10); // Example values: 10, 20, ..., 90
        }

        // Add a scenario named "Baseline"
        ScenarioCollection scenarios = worksheet.Scenarios;
        int scenarioIndex = scenarios.Add("Baseline");
        Scenario baseline = scenarios[scenarioIndex];

        // Capture current values of cells B2 through B10 into the scenario
        for (int row = 1; row <= 9; row++) // Row indices for B2..B10
        {
            // Retrieve the current cell value and convert it to string
            object cellValue = worksheet.Cells[row, 1].Value;
            string valueStr = cellValue != null ? cellValue.ToString() : string.Empty;

            // Add the cell as an input cell to the scenario
            baseline.InputCells.Add(row, 1, valueStr);
        }

        // Save the workbook with the scenario
        workbook.Save("BaselineScenario.xlsx");
    }
}
