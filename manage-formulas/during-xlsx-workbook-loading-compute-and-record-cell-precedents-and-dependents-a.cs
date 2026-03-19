using System;
using System.Collections;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsPrecedentsDependents
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with actual path)
            Workbook workbook = new Workbook("input.xlsx");

            // Enable calculation chain to allow precedents/dependents retrieval
            workbook.Settings.FormulaSettings.EnableCalculationChain = true;

            // Calculate all formulas so the chain is built
            workbook.CalculateFormula();

            // Create a new worksheet to store the trace information
            Worksheet traceSheet = workbook.Worksheets[workbook.Worksheets.Add()];
            traceSheet.Name = "Trace";

            // Write header row
            Cells traceCells = traceSheet.Cells;
            traceCells[0, 0].PutValue("Sheet");
            traceCells[0, 1].PutValue("Cell");
            traceCells[0, 2].PutValue("Precedents (In Calculation)");
            traceCells[0, 3].PutValue("Dependents (Recursive)");

            int outputRow = 1;

            // Iterate through all worksheets
            foreach (Worksheet ws in workbook.Worksheets)
            {
                // Skip the trace sheet itself
                if (ws.Name == traceSheet.Name) continue;

                Cells cells = ws.Cells;
                // Iterate through all cells that contain formulas
                foreach (Cell cell in cells)
                {
                    if (cell.IsFormula)
                    {
                        // ----- Precedents -----
                        StringBuilder precedentsBuilder = new StringBuilder();
                        IEnumerator precedentsEnum = cell.GetPrecedentsInCalculation();
                        if (precedentsEnum != null)
                        {
                            while (precedentsEnum.MoveNext())
                            {
                                ReferredArea area = (ReferredArea)precedentsEnum.Current;
                                // Convert area to a readable string (e.g., Sheet!A1 or Sheet!A1:B2)
                                string areaStr = $"{area.SheetName}!{CellsHelper.CellIndexToName(area.StartRow, area.StartColumn)}";
                                if (area.IsArea)
                                {
                                    areaStr += $":{CellsHelper.CellIndexToName(area.EndRow, area.EndColumn)}";
                                }
                                if (precedentsBuilder.Length > 0) precedentsBuilder.Append(", ");
                                precedentsBuilder.Append(areaStr);
                            }
                        }

                        // ----- Dependents (recursive) -----
                        StringBuilder dependentsBuilder = new StringBuilder();
                        IEnumerator dependentsEnum = cell.GetDependentsInCalculation(true);
                        if (dependentsEnum != null)
                        {
                            while (dependentsEnum.MoveNext())
                            {
                                Cell depCell = (Cell)dependentsEnum.Current;
                                string depStr = $"{depCell.Worksheet.Name}!{depCell.Name}";
                                if (dependentsBuilder.Length > 0) dependentsBuilder.Append(", ");
                                dependentsBuilder.Append(depStr);
                            }
                        }

                        // Write the collected information to the trace sheet
                        traceCells[outputRow, 0].PutValue(ws.Name);
                        traceCells[outputRow, 1].PutValue(cell.Name);
                        traceCells[outputRow, 2].PutValue(precedentsBuilder.ToString());
                        traceCells[outputRow, 3].PutValue(dependentsBuilder.ToString());
                        outputRow++;
                    }
                }
            }

            // Save the workbook with the added trace information
            workbook.Save("output_with_trace.xlsx");
        }
    }
}