using System;
using Aspose.Cells;

namespace AsposeCellsIconSetExample
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range for the icon set (A1:A10)
            CellArea area = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = 9,
                EndColumn = 0
            };

            // Add a new conditional formatting collection
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];
            fcc.AddArea(area);

            // Add an icon set condition
            int conditionIndex = fcc.AddCondition(FormatConditionType.IconSet);
            FormatCondition condition = fcc[conditionIndex];

            // Configure the icon set
            condition.IconSet.Type = IconSetType.TrafficLights31;
            condition.IconSet.ShowValue = true; // Show cell values alongside icons

            // Enable macros in the workbook
            workbook.Settings.EnableMacros = true;

            // Save the workbook as a macro‑enabled file
            string outputPath = "output.xlsm";
            workbook.Save(outputPath, SaveFormat.Xlsm);
        }
    }
}