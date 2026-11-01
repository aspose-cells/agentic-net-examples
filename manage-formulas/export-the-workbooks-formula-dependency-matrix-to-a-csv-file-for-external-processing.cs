// Title: Export an Excel workbook’s formula dependency matrix to a CSV file with Aspose.Cells for .NET (C#)
// Common Searches: c# export excel formula dependency matrix to csv using aspose.cells | how to generate a binary formula map from an xlsx file with Aspose.Cells | aspnet create csv of cell formula presence Aspose.Cells | extract formula flags for each cell and save as csv in C# | Aspose.Cells build formula dependency matrix for first worksheet
// Tags: export formula dependency matrix csv | Aspose.Cells generate formula presence matrix | C# build binary formula map from worksheet | Excel cell.IsFormula extraction to csv | Aspose.Cells used range iteration for formulas

using Aspose.Cells;
using System;
using System.IO;
using System.Text;

// The example loads an Excel workbook with Aspose.Cells, determines the used range of the first worksheet, checks each cell for a formula using the IsFormula property, creates a CSV matrix where '1' marks a formula cell and '0' marks a non‑formula cell, and writes the matrix to 'dependency_matrix.csv', with error handling for missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // For demonstration, use the first worksheet to build a simple dependency matrix
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Determine the used range dimensions
            int rowCount = cells.MaxDataRow + 1;
            int columnCount = cells.MaxDataColumn + 1;

            // Build CSV content: 1 if the cell contains a formula, otherwise 0
            StringBuilder csvBuilder = new StringBuilder();

            for (int i = 0; i < rowCount; i++)
            {
                for (int j = 0; j < columnCount; j++)
                {
                    Cell cell = cells[i, j];
                    bool hasFormula = cell.IsFormula;
                    csvBuilder.Append(hasFormula ? "1" : "0");

                    // Append comma separator except after the last column
                    if (j < columnCount - 1)
                        csvBuilder.Append(",");
                }
                csvBuilder.AppendLine();
            }

            // Save the CSV content to a file
            string outputPath = "dependency_matrix.csv";
            File.WriteAllText(outputPath, csvBuilder.ToString());
            Console.WriteLine($"Dependency matrix saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a message
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
