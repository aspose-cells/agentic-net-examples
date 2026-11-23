// Title: Generate a CSV audit of formula cells with precedent and dependent counts using Aspose.Cells for .NET
// Common Searches: how to list Excel formula cells with their precedent and dependent counts in C# using Aspose.Cells | Aspose.Cells C# export formula dependency statistics to CSV | retrieve precedent count for a formula cell with Aspose.Cells GetPrecedents | get dependent cells count for a formula using Aspose.Cells GetDependents in .NET | create formula audit report from a workbook and save as CSV with Aspose.Cells | C# iterate used range and extract formula information using Aspose.Cells
// Tags: export formula audit to CSV with Aspose.Cells | retrieve formula precedents count using GetPrecedents | retrieve formula dependents count using GetDependents | iterate worksheet used range for formula cells | Aspose.Cells formula dependency report

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The example loads an Excel workbook, walks through each worksheet's used range, identifies formula cells, obtains the number of precedent cells via GetPrecedents and the number of dependent cells via GetDependents(true), and writes these details (sheet name, cell address, formula text, precedent count, dependent count) to a UTF‑8 CSV file while handling missing files and CSV field escaping.
class FormulaAudit
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "formula_audit.csv";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the source workbook
            Workbook workbook;
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Prepare a StringBuilder for CSV content
            StringBuilder csvBuilder = new StringBuilder();

            // Write CSV header
            csvBuilder.AppendLine("Sheet,Cell,Formula,PrecedentCount,DependentCount");

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Determine the used range; skip empty sheets
                int minRow = sheet.Cells.MinRow;
                int maxRow = sheet.Cells.MaxRow;
                int minCol = sheet.Cells.MinColumn;
                int maxCol = sheet.Cells.MaxColumn;

                if (minRow < 0 || minCol < 0) continue; // no used cells

                // Scan each cell in the used range
                for (int row = minRow; row <= maxRow; row++)
                {
                    for (int col = minCol; col <= maxCol; col++)
                    {
                        Cell cell = sheet.Cells[row, col];

                        // Process only formula cells
                        if (cell.IsFormula)
                        {
                            int precedentCount = 0;
                            int dependentCount = 0;

                            try
                            {
                                // Get precedents (cells that this formula depends on)
                                ReferredAreaCollection precedents = cell.GetPrecedents();
                                precedentCount = precedents?.Count ?? 0;
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Error retrieving precedents for {cell.Name}: {ex.Message}");
                            }

                            try
                            {
                                // Get dependents (cells that depend on this formula)
                                // 'true' retrieves all dependents, not only direct ones
                                Cell[] dependents = cell.GetDependents(true);
                                dependentCount = dependents?.Length ?? 0;
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Error retrieving dependents for {cell.Name}: {ex.Message}");
                            }

                            // Build CSV line
                            string line = string.Format(
                                "{0},{1},{2},{3},{4}",
                                EscapeCsv(sheet.Name),
                                EscapeCsv(cell.Name),
                                EscapeCsv(cell.Formula),
                                precedentCount,
                                dependentCount);

                            csvBuilder.AppendLine(line);
                        }
                    }
                }
            }

            // Ensure output directory exists
            try
            {
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the CSV audit file
                File.WriteAllText(outputPath, csvBuilder.ToString(), Encoding.UTF8);
                Console.WriteLine($"Formula audit saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to write CSV file: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }

    // Helper to escape CSV fields that may contain commas or quotes
    private static string EscapeCsv(string field)
    {
        if (string.IsNullOrEmpty(field))
            return string.Empty;

        if (field.Contains(",") || field.Contains("\"") || field.Contains("\n"))
        {
            field = field.Replace("\"", "\"\"");
            return $"\"{field}\"";
        }
        return field;
    }
}
