using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ReplaceTextJoinWithConcat
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook if the file exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates an empty workbook
            }

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all cells that contain formulas
                foreach (Cell cell in sheet.Cells)
                {
                    if (!cell.IsFormula) continue;

                    string formula = cell.Formula;

                    // Look for TEXTJOIN usage (case‑insensitive)
                    if (!Regex.IsMatch(formula, @"TEXTJOIN\s*\(", RegexOptions.IgnoreCase)) continue;

                    // Replace each TEXTJOIN occurrence with an equivalent CONCATENATE formula
                    string newFormula = Regex.Replace(
                        formula,
                        @"TEXTJOIN\s*\(([^)]*)\)",
                        match =>
                        {
                            // Split the arguments of TEXTJOIN
                            string args = match.Groups[1].Value;
                            // Simple split on commas (works for typical cases without nested commas)
                            string[] parts = args.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                            if (parts.Length < 3) return match.Value; // not enough arguments, leave unchanged

                            // First argument: delimiter (keep as is, usually a quoted string)
                            string delimiter = parts[0].Trim();

                            // Remaining arguments are the texts/ranges to join
                            string[] texts = new string[parts.Length - 2];
                            for (int i = 2; i < parts.Length; i++)
                                texts[i - 2] = parts[i].Trim();

                            // Build CONCATENATE argument list interleaving the delimiter
                            // Example: CONCATENATE(text1, delimiter, text2, delimiter, text3)
                            string concatArgs = texts[0];
                            for (int i = 1; i < texts.Length; i++)
                            {
                                concatArgs += $", {delimiter}, {texts[i]}";
                            }

                            // Return the new formula fragment
                            return $"CONCATENATE({concatArgs})";
                        },
                        RegexOptions.IgnoreCase);

                    // Update the cell with the transformed formula
                    cell.Formula = newFormula;
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}