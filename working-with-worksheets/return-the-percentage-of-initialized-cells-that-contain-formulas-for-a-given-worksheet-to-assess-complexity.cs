// Title: Calculate the percentage of initialized cells that contain formulas in an Excel worksheet using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells get percentage of formula cells in a worksheet | how to compute formula density of an Excel sheet with Aspose.Cells | determine proportion of formula cells versus data cells using Aspose.Cells .NET | Aspose.Cells count initialized cells and formula cells in a worksheet | measure spreadsheet complexity by formula cell ratio in C#
// Tags: calculate formula cell ratio Aspose.Cells | worksheet formula density analysis .NET | count initialized cells Aspose.Cells | used range iteration Aspose.Cells C# | measure spreadsheet formula complexity

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsFormulaComplexity
{
    // The GetFormulaCellPercentage method loads a workbook, selects the specified worksheet, determines its used range via MaxDataRow/MaxDataColumn, iterates through each cell counting initialized cells (non‑null value or formula) and formula cells, then returns the formula‑cell count divided by the initialized‑cell count as a percentage (0‑100). It handles empty sheets, missing files, and invalid worksheet names with appropriate exceptions.
    public class WorksheetFormulaAnalyzer
    {
        /// <param name="workbookPath">Full path to the Excel file.</param>
        /// <param name="worksheetName">Name of the worksheet to analyze.</param>
        /// <returns>Percentage of initialized cells that are formulas (0‑100). Returns 0 if there are no initialized cells.</returns>
        public static double GetFormulaCellPercentage(string workbookPath, string worksheetName)
        {
            // Verify that the file exists to avoid FileNotFoundException
            if (!File.Exists(workbookPath))
                throw new FileNotFoundException($"Workbook file not found: {workbookPath}");

            Workbook workbook;
            try
            {
                // Load the workbook
                workbook = new Workbook(workbookPath);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to load workbook: {ex.Message}", ex);
            }

            // Get the target worksheet
            Worksheet sheet = workbook.Worksheets[worksheetName];
            if (sheet == null)
                throw new ArgumentException($"Worksheet \"{worksheetName}\" not found.");

            // Determine the used range (max row/column with data)
            int maxRow = sheet.Cells.MaxDataRow;
            int maxCol = sheet.Cells.MaxDataColumn;

            // If the sheet is completely empty, return 0
            if (maxRow < 0 || maxCol < 0)
                return 0.0;

            int initializedCellCount = 0;
            int formulaCellCount = 0;

            // Iterate through the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = sheet.Cells[row, col];

                    // Consider a cell initialized if it contains a formula or a non‑null value
                    bool isInitialized = cell.IsFormula || cell.Value != null;
                    if (!isInitialized)
                        continue;

                    initializedCellCount++;

                    // Count cells that contain a formula
                    if (cell.IsFormula)
                        formulaCellCount++;
                }
            }

            // Avoid division by zero
            if (initializedCellCount == 0)
                return 0.0;

            // Calculate percentage
            double percentage = (double)formulaCellCount / initializedCellCount * 100.0;
            return percentage;
        }

        // Example usage
        public static void Main()
        {
            try
            {
                string filePath = @"C:\Data\Sample.xlsx";
                string sheetName = "Sheet1";

                double formulaPercentage = GetFormulaCellPercentage(filePath, sheetName);
                Console.WriteLine($"Formula cells: {formulaPercentage:F2}% of initialized cells in \"{sheetName}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
