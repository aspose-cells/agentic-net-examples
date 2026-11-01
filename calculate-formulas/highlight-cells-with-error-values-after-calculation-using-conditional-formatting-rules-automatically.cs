// Title: Automatically highlight cells with formula errors using Aspose.Cells conditional formatting in C#
// Common Searches: how to use Aspose.Cells in C# to apply conditional formatting that flags #DIV/0! errors | C# code example for highlighting Excel cells containing errors with Aspose.Cells | apply red background to error cells after CalculateFormula with Aspose.Cells .NET | Aspose.Cells conditional formatting for cells that contain errors in a workbook | automate error detection and highlighting in Excel using Aspose.Cells for .NET
// Tags: Aspose.Cells error cell highlighting | C# red fill for formula errors | calculate formulas before applying style Aspose.Cells | Excel workbook error detection Aspose.Cells .NET | apply style to cells containing errors Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example loads an existing XLSX file, forces formula recalculation, defines the used range, adds a formatting rule that detects cells with errors, applies a red solid background to those cells, and saves the workbook, demonstrating automatic error highlighting with Aspose.Cells for .NET.
class HighlightErrorCells
{
    static void Main()
    {
        try
        {
            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Ensure all formulas are calculated so that error values are up‑to‑date
            workbook.CalculateFormula();

            // Get the first worksheet (or specify the desired one)
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range to which the conditional formatting will be applied.
            // Here we apply it to the entire used range of the worksheet.
            CellArea area = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = sheet.Cells.MaxDataRow,
                EndColumn = sheet.Cells.MaxDataColumn
            };

            // Add a conditional formatting collection to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection collection = sheet.ConditionalFormattings[cfIndex];
            collection.AddArea(area);

            // Create a condition that checks for cells containing errors
            int conditionIndex = collection.AddCondition(FormatConditionType.ContainsErrors);
            FormatCondition condition = collection[conditionIndex];

            // Define the style to apply when the condition is met (e.g., red fill)
            Style style = workbook.CreateStyle();
            style.ForegroundColor = Color.Red;
            style.Pattern = BackgroundType.Solid;
            condition.Style = style;

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
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
