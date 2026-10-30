// Title: Load an XLS workbook, add a greater‑than conditional formatting rule to A1:A10, and export it as an MHT web archive using Aspose.Cells for .NET (C#)
// Common Searches: how to apply conditional formatting to a range in an existing XLS workbook using Aspose.Cells C# | save a workbook as MHT web archive after formatting cells with Aspose.Cells .NET | convert XLS to MHT while preserving conditional formatting with Aspose.Cells | C# example for adding a greater‑than condition and exporting to MHT using Aspose.Cells | load XLS, set cell background based on value, and export as .mht in C#
// Tags: Aspose.Cells add cell value condition C# | Aspose.Cells style cells based on value | Aspose.Cells export workbook to MHT | Aspose.Cells Html save format for web archive | C# load XLS and apply conditional formatting

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;

// The sample loads an existing XLS file, creates a conditional formatting rule that highlights cells A1:A10 with a light coral background when their value exceeds 50, and then saves the modified workbook as an MHT web archive using the Html save format. It includes a file‑existence check and basic exception handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xls";
            const string outputPath = "output.mht";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"The input file '{inputPath}' was not found.");

            // Load the existing XLS file
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range A1:A10 for conditional formatting
            CellArea area = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = 9,
                EndColumn = 0
            };

            // Add a new conditional formatting collection to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];
            fcc.AddArea(area);

            // Create a condition: cell value greater than 50
            FormatCondition condition = fcc[0];
            condition.Type = FormatConditionType.CellValue;
            condition.Operator = OperatorType.GreaterThan;
            condition.Formula1 = "50";

            // Define the style to apply when the condition is met (light red background)
            Style style = workbook.CreateStyle();
            style.ForegroundColor = Color.LightCoral;
            style.Pattern = BackgroundType.Solid;
            condition.Style = style;

            // Save the workbook as an MHT web archive (using Html format with .mht extension)
            workbook.Save(outputPath, SaveFormat.Html);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
