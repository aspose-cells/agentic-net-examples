using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Cells;

namespace AsposeCellsRefValidation
{
    class Program
    {
        static void Main()
        {
            // ---------- Create a new workbook (target) ----------
            Workbook targetWorkbook = new Workbook();

            // ---------- Load source workbooks ----------
            // Replace with actual file paths
            string sourcePath1 = "Source1.xlsx";
            string sourcePath2 = "Source2.xlsx";

            Workbook sourceWorkbook1 = new Workbook(sourcePath1);
            Workbook sourceWorkbook2 = new Workbook(sourcePath2);

            // ---------- Merge worksheets from source workbooks into target ----------
            MergeWorksheets(sourceWorkbook1, targetWorkbook);
            MergeWorksheets(sourceWorkbook2, targetWorkbook);

            // ---------- Calculate formulas without ignoring errors ----------
            CalculationOptions calcOptions = new CalculationOptions
            {
                IgnoreError = false
            };
            targetWorkbook.CalculateFormula(calcOptions);

            // ---------- Validate that no #REF! errors exist ----------
            List<string> refErrorCells = new List<string>();

            foreach (Worksheet sheet in targetWorkbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];
                        if (cell.Type == CellValueType.IsError && cell.StringValue == "#REF!")
                        {
                            refErrorCells.Add($"{sheet.Name}!{cell.Name}");
                        }
                    }
                }
            }

            // ---------- Report results ----------
            if (refErrorCells.Count == 0)
            {
                Console.WriteLine("Validation passed: No #REF! errors found in any cell.");
            }
            else
            {
                Console.WriteLine("Validation failed: #REF! errors detected in the following cells:");
                foreach (string address in refErrorCells)
                {
                    Console.WriteLine(address);
                }
            }

            // ---------- Save the merged workbook ----------
            // Replace with desired output path
            string outputPath = "MergedOutput.xlsx";
            targetWorkbook.Save(outputPath);
            Console.WriteLine($"Merged workbook saved to: {outputPath}");
        }

        private static void MergeWorksheets(Workbook source, Workbook target)
        {
            foreach (Worksheet srcSheet in source.Worksheets)
            {
                // Add a new empty worksheet to the target workbook
                int newIndex = target.Worksheets.Add();
                Worksheet newSheet = target.Worksheets[newIndex];

                // Ensure a unique name
                string baseName = srcSheet.Name;
                string uniqueName = GetUniqueWorksheetName(target, baseName);
                newSheet.Name = uniqueName;

                // Copy contents
                srcSheet.Copy(newSheet);
            }
        }

        private static string GetUniqueWorksheetName(Workbook workbook, string baseName)
        {
            string name = baseName;
            int suffix = 1;
            while (workbook.Worksheets.Any(ws => ws.Name.Equals(name, StringComparison.OrdinalIgnoreCase)))
            {
                name = $"{baseName}_{suffix}";
                suffix++;
            }
            return name;
        }
    }
}