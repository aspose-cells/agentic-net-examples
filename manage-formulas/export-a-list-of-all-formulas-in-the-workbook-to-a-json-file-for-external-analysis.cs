// Title: Export all worksheet formulas to a JSON file using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an Excel workbook with Aspose.Cells, iterates every worksheet, extracts each cell's formula together with its sheet name and address, and writes the collection as a formatted JSON file. | Update the sample program so that it ignores hidden worksheets while exporting formulas to JSON with Aspose.Cells. | Enhance the formula extraction routine by adding console logging that reports the number of formulas processed for each worksheet.
// Common Searches: how to use Aspose.Cells in C# to export cell formulas to a JSON document | C# Aspose.Cells extract formulas from all sheets and save as JSON | export Excel formulas with sheet name and cell reference to JSON using Aspose.Cells | Aspose.Cells .NET list formulas and serialize to JSON file
// Tags: Aspose.Cells serialize formulas as JSON | C# enumerate worksheet formulas Aspose.Cells | JsonSerializer write formula metadata .NET | filter out hidden worksheets Aspose.Cells | log formula counts per worksheet Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace ExportFormulasToJson
{
    // DTO to hold formula information
    // The example loads an Excel workbook via Aspose.Cells, walks through each worksheet and its cells, captures cells that contain formulas along with their sheet names and addresses into a list of FormulaInfo objects, serializes this list to an indented JSON string using System.Text.Json, ensures the target directory exists, and writes the JSON to the specified output path.
    public class FormulaInfo
    {
        public string SheetName { get; set; }
        public string CellName { get; set; }
        public string Formula { get; set; }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Expect two arguments: input workbook path and output JSON path
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: ExportFormulasToJson <InputWorkbookPath> <OutputJsonPath>");
                return;
            }

            string inputPath = args[0];
            string outputPath = args[1];

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Collect formulas from all worksheets
                List<FormulaInfo> formulas = new List<FormulaInfo>();

                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;

                    // Iterate through all cells that contain formulas
                    foreach (Cell cell in cells)
                    {
                        if (cell.IsFormula && !string.IsNullOrEmpty(cell.Formula))
                        {
                            formulas.Add(new FormulaInfo
                            {
                                SheetName = sheet.Name,
                                CellName = cell.Name,
                                Formula = cell.Formula
                            });
                        }
                    }
                }

                // Serialize to JSON with indentation
                string json = JsonSerializer.Serialize(formulas, new JsonSerializerOptions { WriteIndented = true });

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Write JSON to file
                File.WriteAllText(outputPath, json);

                Console.WriteLine($"Successfully exported {formulas.Count} formulas to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
