// Title: Apply red background conditional formatting to cells A1:A20 and export the workbook as HTML using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# apply conditional formatting to a specific range before saving as HTML | How to set a red background for cells with values greater than 100 using Aspose.Cells and generate HTML output | C# example adding a cell value > 100 conditional format and preserving it in HTML export with Aspose.Cells | Export Excel to HTML while retaining conditional formatting using Aspose.Cells for .NET | Aspose.Cells conditional formatting range A1:A20 red fill HTML conversion sample
// Tags: conditional formatting range Aspose.Cells | red background style Aspose.Cells | HTML export preserving conditional formatting Aspose.Cells | cell value greater than 100 format Aspose.Cells | C# workbook to HTML Aspose.Cells

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace ConditionalFormattingToHtml
{
    // The program loads an existing XLSX file, creates a conditional formatting rule for the range A1:A20 that highlights cells with values greater than 100 using a solid red background, and then saves the workbook as an HTML file while retaining the visual formatting.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Define the range for conditional formatting (A1:A20)
                int firstRow = 0;      // zero‑based index
                int firstColumn = 0;   // zero‑based index
                int totalRows = 20;
                int totalColumns = 1;
                AsposeRange range = sheet.Cells.CreateRange(firstRow, firstColumn, totalRows, totalColumns);

                // Add a new conditional formatting collection to the worksheet
                int cfIndex = sheet.ConditionalFormattings.Add();
                FormatConditionCollection collection = sheet.ConditionalFormattings[cfIndex];

                // Convert Range to CellArea (required by AddArea)
                CellArea area = new CellArea
                {
                    StartRow = firstRow,
                    StartColumn = firstColumn,
                    EndRow = firstRow + totalRows - 1,
                    EndColumn = firstColumn + totalColumns - 1
                };
                collection.AddArea(area);

                // Create a condition: Cell Value > 100
                FormatCondition condition = collection[collection.Count - 1];
                condition.Type = FormatConditionType.CellValue;
                condition.Operator = OperatorType.GreaterThan;
                condition.Formula1 = "100";

                // Define the formatting style (red background)
                Style style = workbook.CreateStyle();
                style.ForegroundColor = Color.Red;
                style.Pattern = BackgroundType.Solid;
                condition.Style = style;

                // Save the workbook as HTML, preserving conditional formatting appearance
                workbook.Save(outputPath, SaveFormat.Html);
                Console.WriteLine($"Workbook saved as HTML to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
