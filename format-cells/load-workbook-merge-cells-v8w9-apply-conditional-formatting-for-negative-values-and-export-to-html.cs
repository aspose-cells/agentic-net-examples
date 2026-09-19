// Title: Merge cells V8:W9, highlight negative values in red, and save the worksheet as HTML with Aspose.Cells for .NET
// AI Prompts: Create a C# program that opens an Excel workbook, merges the range V8:W9, adds a conditional formatting rule to display values less than zero in red, and exports the result to an HTML file using Aspose.Cells. | Write Aspose.Cells for .NET code to combine cells V8 through W9, apply a red‑font style to any negative numbers via conditional formatting, and save the workbook as HTML.
// Common Searches: Aspose.Cells C# merge V8 W9 cells and apply red font conditional formatting for values < 0 | Export Excel to HTML with merged cells and negative number highlighting using Aspose.Cells | How to add conditional formatting for negative numbers in an Aspose.Cells workbook before saving as HTML | C# example for merging a specific cell range and setting conditional format then converting to HTML with Aspose.Cells
// Tags: merge specific cell range Aspose.Cells C# | negative value conditional formatting Aspose.Cells | save workbook as HTML Aspose.Cells | apply red font style Aspose.Cells | cell area V8:W9 Aspose.Cells

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Loads an existing Excel file (or creates a new workbook), merges cells V8:W9, adds a conditional formatting rule that colors any value less than zero red, and saves the workbook as an HTML document using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Define input and output file paths
                string inputPath = "input.xlsx";
                string outputPath = "output.html";

                // Load an existing workbook if the file exists; otherwise create a new one
                Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Merge cells V8:W9 (zero‑based indexes: row 7‑8, column 21‑22)
                sheet.Cells.Merge(7, 21, 2, 2);

                // Define the range for conditional formatting (entire used range)
                CellArea formatRange = CellArea.CreateCellArea(
                    0, 0,
                    sheet.Cells.MaxDataRow,
                    sheet.Cells.MaxDataColumn);

                // Add a new conditional formatting collection to the worksheet
                int cfIndex = sheet.ConditionalFormattings.Add();
                FormatConditionCollection cfCollection = sheet.ConditionalFormattings[cfIndex];
                cfCollection.AddArea(formatRange);

                // Add a condition: cell value < 0
                int conditionIndex = cfCollection.AddCondition(
                    FormatConditionType.CellValue,
                    OperatorType.LessThan,
                    "0",
                    null);
                FormatCondition condition = cfCollection[conditionIndex];

                // Define style for negative values (red font)
                Style style = workbook.CreateStyle();
                style.Font.Color = Color.Red;
                condition.Style = style;

                // Save the workbook as HTML
                HtmlSaveOptions saveOptions = new HtmlSaveOptions();
                workbook.Save(outputPath, saveOptions);

                Console.WriteLine($"Workbook processed successfully. Output saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
