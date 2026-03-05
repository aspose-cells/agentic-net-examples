using System;
using Aspose.Cells;

namespace AsposeCellsLocalizationDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Create a SettableGlobalizationSettings instance to customize labels
            SettableGlobalizationSettings globalization = new SettableGlobalizationSettings();

            // Localize total names for different consolidation functions (example: French)
            globalization.SetTotalName(ConsolidationFunction.Sum, "Somme");
            globalization.SetTotalName(ConsolidationFunction.Count, "Nombre");
            globalization.SetTotalName(ConsolidationFunction.Average, "Moyenne");
            globalization.SetTotalName(ConsolidationFunction.Max, "Maximum");
            globalization.SetTotalName(ConsolidationFunction.Min, "Minimum");

            // Localize the grand total label
            globalization.SetGrandTotalName(ConsolidationFunction.Sum, "Total Général");
            globalization.SetGrandTotalName(ConsolidationFunction.Count, "Nombre Total");
            globalization.SetGrandTotalName(ConsolidationFunction.Average, "Moyenne Totale");
            globalization.SetGrandTotalName(ConsolidationFunction.Max, "Maximum Total");
            globalization.SetGrandTotalName(ConsolidationFunction.Min, "Minimum Total");

            // Apply the custom globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = globalization;

            // Define the range on which to apply subtotals (example: first sheet, columns A and B)
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;
            CellArea area = CellArea.CreateCellArea(0, 0, cells.MaxDataRow, 1); // rows 0..max, columns A and B

            // Apply subtotal: group by column 0 (A), use Sum on column 1 (B), include subtotals and grand total
            cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 0 }, true, false, true);

            // Save the modified workbook
            string outputPath = "output_localized.xlsx";
            workbook.Save(outputPath);
        }
    }
}