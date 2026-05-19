using Aspose.Cells;
using System;
using System.IO;
using System.Text;

// Alias to avoid conflict with System.Range (C# 8+)
using AsposeRange = Aspose.Cells.Range;

class ExportFormulasToLaTeX
{
    static void Main()
    {
        try
        {
            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output_with_latex.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Collect LaTeX representations of all formulas
            StringBuilder latexBuilder = new StringBuilder();

            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Determine the used range of the worksheet
                AsposeRange usedRange = cells.MaxDisplayRange;
                if (usedRange == null)
                    continue; // Skip empty sheets

                int startRow = usedRange.FirstRow;
                int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
                int startCol = usedRange.FirstColumn;
                int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                for (int row = startRow; row <= endRow; row++)
                {
                    for (int col = startCol; col <= endCol; col++)
                    {
                        Cell cell = cells[row, col];
                        if (!string.IsNullOrEmpty(cell.Formula))
                        {
                            // Simple conversion: strip leading '=' and wrap with LaTeX math delimiters
                            string formulaBody = cell.Formula.TrimStart('=');
                            string latex = "$" + formulaBody + "$";
                            latexBuilder.AppendLine(latex);
                        }
                    }
                }
            }

            // Add a LaTeX equation shape to the first worksheet containing all formulas
            Worksheet firstSheet = workbook.Worksheets[0];
            firstSheet.Shapes.AddLaTeXEquation(
                topRow: 0,
                top: 0,
                leftColumn: 0,
                left: 0,
                height: 200,
                width: 400,
                latex: latexBuilder.ToString());

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}