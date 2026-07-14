using System;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source workbook (must exist)
            string sourcePath = "input.xlsx";

            // Load the workbook using the string constructor (provided rule)
            using (Workbook workbook = new Workbook(sourcePath))
            {
                // Modify default style
                workbook.DefaultStyle.Font.Name = "Arial";
                workbook.DefaultStyle.Font.Size = 12;

                // Update built‑in document property
                workbook.BuiltInDocumentProperties["Author"].Value = "John Doe";

                // Add a custom document property
                workbook.CustomDocumentProperties.Add("Reviewed", true);

                // Enable iterative calculation in workbook settings
                workbook.Settings.FormulaSettings.EnableIterativeCalculation = true;
                workbook.Settings.FormulaSettings.MaxIteration = 100;
                workbook.Settings.FormulaSettings.MaxChange = 0.001;

                // Save the modified workbook to XLSX format (provided Save rule)
                workbook.Save("output.xlsx", SaveFormat.Xlsx);
            }
        }
    }
}