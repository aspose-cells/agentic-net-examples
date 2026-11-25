// Title: How to create a Baseline scenario that captures values of cells B2‑B10 with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add scenario named Baseline for range B2:B10 | capture current worksheet cell values into ScenarioInputCellCollection using Aspose.Cells | set active scenario index after adding a scenario in Aspose.Cells worksheet | save workbook with scenario data to XLSX using Aspose.Cells C# | populate cells programmatically then create a baseline scenario with Aspose.Cells
// Tags: Aspose.Cells scenario collection add baseline | capture worksheet range B2:B10 into scenario input cells | set active scenario index Aspose.Cells .NET | save workbook with scenarios to XLSX | populate cells programmatically Aspose.Cells C#

using System;
using Aspose.Cells;

namespace ScenarioBaselineDemo
{
    // The example creates a new workbook, fills cells B2 through B10 with sample data, adds a scenario named "Baseline" to the worksheet's ScenarioCollection, records the current values of those cells as input cells, sets the new scenario as the active one, and saves the workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the output workbook
            string outputPath = "BaselineScenario.xlsx";

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Name = "Data";

            // Populate cells B2 through B10 with sample data
            for (int row = 1; row <= 9; row++) // B2 is row index 1 (zero‑based)
            {
                // Example: put the row number as the value
                worksheet.Cells[row, 1].PutValue($"Value_{row + 1}");
            }

            // Access the scenario collection of the worksheet
            ScenarioCollection scenarios = worksheet.Scenarios;

            // Add a new scenario named "Baseline"
            int scenarioIndex = scenarios.Add("Baseline");
            Scenario baselineScenario = scenarios[scenarioIndex];

            // Capture current values of cells B2:B10 into the scenario's input cells
            ScenarioInputCellCollection inputCells = baselineScenario.InputCells;

            for (int row = 1; row <= 9; row++) // rows 1..9 correspond to B2..B10
            {
                // Retrieve the current value as string (handles numbers, strings, dates, etc.)
                string cellValue = worksheet.Cells[row, 1].StringValue;

                // Add the input cell to the scenario (row, column, value)
                // Column index for B is 1 (zero‑based)
                inputCells.Add(row, 1, cellValue);
            }

            // Optionally set the newly created scenario as the active one
            scenarios.ActiveIndex = scenarioIndex;
            scenarios.LastSelected = scenarioIndex;

            // Save the workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved to '{outputPath}'. Scenario 'Baseline' created for cells B2:B10.");
        }
    }
}
