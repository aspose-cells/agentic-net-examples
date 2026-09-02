// Title: Calculate the percentage of initialized cells that contain formulas in an Aspose.Cells worksheet using C#
// AI Prompts: Write a C# method that uses Aspose.Cells to return the ratio of formula cells to all initialized (non‑empty) cells in a worksheet. | Update the sample to ignore cells that only have formatting when computing the formula percentage. | Build a console program that loads an Excel file with Aspose.Cells, calls the formula‑percentage method, and prints the result with two decimal places. | Add robust error handling to the GetFormulaCellPercentage function to safely handle worksheets with zero initialized cells.
// Common Searches: aspnet cells get percentage of cells with formulas in a worksheet | c# aspose.cells calculate formula cell ratio for excel file | how to measure worksheet complexity by counting formula cells using Aspose.Cells | determine formula cell count versus total initialized cells in Aspose.Cells C# | skip empty and formatting‑only cells when computing formula percentage Aspose.Cells
// Tags: calculate formula cell percentage Aspose.Cells C# | initialized cell count with formulas Aspose.Cells | worksheet complexity metric formula ratio | iterate worksheet cells Aspose.Cells | exclude formatting‑only cells Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example defines GetFormulaCellPercentage, which iterates through every cell in a given Aspose.Cells Worksheet, counts cells that are initialized (contain a value or a formula), counts how many of those contain formulas, and returns the formula count as a percentage of initialized cells. The Main method loads a workbook, selects the first worksheet, invokes the method, and prints the percentage.
public class WorksheetAnalysis
{
    /// <param name="worksheet">The worksheet to analyze.</param>
    /// <returns>The percentage of formula cells.</returns>
    public static double GetFormulaCellPercentage(Worksheet worksheet)
    {
        int totalInitialized = 0;
        int formulaCount = 0;

        // Iterate through all cells in the worksheet
        foreach (Cell cell in worksheet.Cells)
        {
            // Consider a cell initialized if it has a value or a formula
            bool hasValue = cell.Value != null;
            bool hasFormula = !string.IsNullOrEmpty(cell.Formula);

            if (!hasValue && !hasFormula)
                continue; // skip truly empty cells

            totalInitialized++;

            if (hasFormula)
                formulaCount++;
        }

        if (totalInitialized == 0)
            return 0.0;

        return (double)formulaCount / totalInitialized * 100.0;
    }

    /// <summary>
    /// Entry point for demonstration.
    /// </summary>
    public static void Main()
    {
        try
        {
            string filePath = "Sample.xlsx";

            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Load workbook
            Workbook workbook = new Workbook(filePath);
            Worksheet sheet = workbook.Worksheets[0];

            double percentage = GetFormulaCellPercentage(sheet);
            Console.WriteLine($"Formula cell percentage: {percentage:F2}%");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
