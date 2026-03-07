using System;
using Aspose.Cells;

namespace AsposeCellsScenarioDemo
{
    public class ScenarioUsage
    {
        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some cells that will be used as input cells for scenarios
            worksheet.Cells["A1"].PutValue(10);
            worksheet.Cells["A2"].PutValue(20);
            worksheet.Cells["B1"].PutValue(30);
            worksheet.Cells["B2"].PutValue(40);

            // Access the scenario collection of the worksheet
            ScenarioCollection scenarios = worksheet.Scenarios;

            // Add first scenario and configure it
            int index1 = scenarios.Add("BaseScenario");
            Scenario scenario1 = scenarios[index1];
            scenario1.Comment = "Base scenario with original values";
            scenario1.IsHidden = false;
            scenario1.IsLocked = true;
            // Define input cells for the first scenario
            scenario1.InputCells.Add(0, 0, "10"); // A1
            scenario1.InputCells.Add(1, 0, "20"); // A2

            // Add second scenario and configure it
            int index2 = scenarios.Add("AdjustedScenario");
            Scenario scenario2 = scenarios[index2];
            scenario2.Comment = "Adjusted values for testing";
            scenario2.IsHidden = false;
            scenario2.IsLocked = false;
            // Define input cells for the second scenario
            scenario2.InputCells.Add(0, 1, "35"); // B1
            scenario2.InputCells.Add(1, 1, "45"); // B2

            // Set active scenario and last selected scenario indices
            scenarios.ActiveIndex = index2;      // Make AdjustedScenario the active one
            scenarios.LastSelected = index2;    // Record it as the last selected

            // Protect the worksheet but allow editing of scenarios
            Protection protection = worksheet.Protection;
            protection.AllowEditingScenario = true;
            protection.Password = "secure123";
            worksheet.Protect(ProtectionType.All);

            // Save the workbook with scenarios
            workbook.Save("ScenarioUsageDemo.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ScenarioUsage.Run();
        }
    }
}