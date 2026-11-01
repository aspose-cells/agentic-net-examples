// Title: How to recalculate conditional formatting formulas after applying a rule using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# recalculate conditional formatting after adding a rule | force conditional formatting to refresh in an Aspose.Cells workbook | call Workbook.CalculateFormula to update conditional formatting colors | programmatically apply conditional formatting and then recalc formulas with Aspose.Cells | refresh conditional formatting after changing cell values using Aspose.Cells .NET
// Tags: conditional formatting formula recalculation Aspose.Cells | Workbook.CalculateFormula method C# | apply conditional formatting Aspose.Cells .NET | refresh conditional formatting after data change | Aspose.Cells Excel conditional formatting example

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsConditionalFormattingExample
{
    // The example creates a new workbook, fills cells A1‑A3 with values, adds a conditional formatting rule that highlights cells greater than 15 in red, invokes Workbook.CalculateFormula to recalculate any formulas used by the conditional formatting, and saves the file as ConditionalFormatting.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate some sample data
                sheet.Cells["A1"].PutValue(10);
                sheet.Cells["A2"].PutValue(20);
                sheet.Cells["A3"].PutValue(30);

                // Add a new conditional formatting collection to the worksheet
                int cfIndex = sheet.ConditionalFormattings.Add();
                FormatConditionCollection cfCollection = sheet.ConditionalFormattings[cfIndex];

                // Add a condition: cells with values greater than 15
                int conditionIndex = cfCollection.AddCondition(
                    FormatConditionType.CellValue,
                    OperatorType.GreaterThan,
                    "15",
                    string.Empty);

                // Retrieve the created condition
                FormatCondition cf = cfCollection[conditionIndex];

                // Define the style for the condition (red background)
                Style style = workbook.CreateStyle();
                style.ForegroundColor = Color.Red;
                style.Pattern = BackgroundType.Solid;
                cf.Style = style;

                // Recalculate all formulas (including those used by conditional formatting)
                workbook.CalculateFormula();

                // Prepare output path
                string outputFile = "ConditionalFormatting.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputFile));

                // Ensure the directory exists (outputDir can be null if only a file name is provided)
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
