using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsAudit
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Enable calculation chain and calculate all formulas
            workbook.Settings.FormulaSettings.EnableCalculationChain = true;
            workbook.CalculateFormula();

            // Save the workbook after calculation (lifecycle rule)
            workbook.Save("input_calculated.xlsx");

            // Prepare CSV output
            using (StreamWriter csvWriter = new StreamWriter("audit.csv"))
            {
                // Write CSV header
                csvWriter.WriteLine("FormulaCell,Precedents,DependentCount");

                // Process each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;

                    // Determine the used range
                    int maxRow = cells.MaxDataRow;
                    int maxCol = cells.MaxDataColumn;

                    // Iterate through all cells in the used range
                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxCol; col++)
                        {
                            Cell cell = cells[row, col];

                            // Consider only formula cells
                            if (cell.IsFormula)
                            {
                                // ----- Collect precedents -----
                                string precedentsText = "";
                                IEnumerator precedentsEnum = cell.GetPrecedentsInCalculation();
                                if (precedentsEnum != null)
                                {
                                    var precedentsList = new System.Collections.Generic.List<string>();
                                    while (precedentsEnum.MoveNext())
                                    {
                                        // Each item is a ReferredArea
                                        if (precedentsEnum.Current is ReferredArea area)
                                        {
                                            // Use the area string representation (e.g., Sheet1!A1:B2)
                                            precedentsList.Add(area.ToString());
                                        }
                                    }
                                    precedentsText = string.Join(";", precedentsList);
                                }

                                // ----- Count dependents -----
                                int dependentCount = 0;
                                IEnumerator dependentsEnum = cell.GetDependentsInCalculation(true);
                                if (dependentsEnum != null)
                                {
                                    while (dependentsEnum.MoveNext())
                                    {
                                        // Each item is a Cell
                                        if (dependentsEnum.Current is Cell)
                                        {
                                            dependentCount++;
                                        }
                                    }
                                }

                                // Write CSV line
                                csvWriter.WriteLine($"{cell.Name},\"{precedentsText}\",{dependentCount}");
                            }
                        }
                    }
                }
            }

            Console.WriteLine("Audit CSV file 'audit.csv' has been generated.");
        }
    }
}