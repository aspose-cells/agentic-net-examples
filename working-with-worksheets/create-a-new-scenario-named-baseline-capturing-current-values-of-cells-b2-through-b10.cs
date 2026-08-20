// Title: Add a 'Baseline' scenario that captures B2:B10 values using Aspose.Cells for .NET
// Description: Demonstrates how to create a new workbook, optionally fill cells B2‑B10, add a scenario named "Baseline", record the current values of those cells as input cells, and save the file as BaselineScenario.xlsx with Aspose.Cells for C#.
// Keywords: Aspose.Cells scenario | C# add scenario | capture cell values | Baseline scenario | input cells B2 B10 | save workbook Aspose.Cells | .NET Excel scenario API
// Common Searches: Aspose.Cells create scenario Baseline | how to store B2:B10 values in a scenario .NET | add input cells to Aspose.Cells scenario | record current cell values in Excel scenario using C# | save workbook with scenario Aspose.Cells
// Developer Intent: Create a scenario called Baseline and store the existing values of cells B2 through B10 as input cells.
// Use Cases: Establish a reference data set before performing what‑if analysis on financial models. | Provide users with a default state that can be restored after editing worksheet values. | Generate a template where the baseline scenario serves as a comparison point for multiple outcome scenarios.
// AI Prompts: Show code to add additional scenarios and switch between them programmatically with Aspose.Cells for .NET. | Provide an example that updates the Baseline scenario after modifying values in B2:B10. | Explain how to export all scenario input cells to a separate worksheet for reporting.

using System;
using Aspose.Cells;

namespace AsposeCellsScenarioDemo
{
    // Demonstrates how to create a new workbook, optionally fill cells B2‑B10, add a scenario named "Baseline", record the current values of those cells as input cells, and save the file as BaselineScenario.xlsx with Aspose.Cells for C#.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // (Optional) Populate cells B2:B10 with sample data
            for (int row = 1; row <= 9; row++) // B2 is row index 1, B10 is row index 9
            {
                worksheet.Cells[row, 1].PutValue($"Value{row}");
            }

            // Add a new scenario named "Baseline"
            int scenarioIndex = worksheet.Scenarios.Add("Baseline");
            Scenario baselineScenario = worksheet.Scenarios[scenarioIndex];

            // Capture current values of cells B2 through B10 into the scenario
            for (int row = 1; row <= 9; row++) // rows 1 to 9 correspond to B2:B10
            {
                // Retrieve the cell's current value as a string
                string cellValue = worksheet.Cells[row, 1].Value?.ToString() ?? string.Empty;

                // Add the cell as an input cell to the scenario
                baselineScenario.InputCells.Add(row, 1, cellValue);
            }

            // Save the workbook with the scenario
            workbook.Save("BaselineScenario.xlsx");
        }
    }
}
