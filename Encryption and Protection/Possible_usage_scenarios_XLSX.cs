using System;
using Aspose.Cells;

namespace AsposeCellsScenariosDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (uses Workbook() constructor)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add sample data to cells
            worksheet.Cells["A1"].PutValue("Initial Value");
            worksheet.Cells["A2"].PutValue(100);

            // Get the ScenarioCollection for the worksheet
            ScenarioCollection scenarios = worksheet.Scenarios;

            // Add first scenario and set its comment
            int scenarioIndex1 = scenarios.Add("Scenario1");
            Scenario scenario1 = scenarios[scenarioIndex1];
            scenario1.Comment = "First test scenario";

            // Add second scenario and set its comment
            int scenarioIndex2 = scenarios.Add("Scenario2");
            Scenario scenario2 = scenarios[scenarioIndex2];
            scenario2.Comment = "Second test scenario";

            // Save the workbook with scenarios (uses Workbook.Save(string) method)
            workbook.Save("WorksheetScenariosDemo.xlsx");

            // Clear all scenarios from the collection
            scenarios.Clear();

            // Save the workbook after clearing scenarios
            workbook.Save("WorksheetScenariosDemo_Cleared.xlsx");
        }
    }
}