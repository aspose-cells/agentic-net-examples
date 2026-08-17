// Title: Apply Brazilian Portuguese Custom Percentage Format to Excel Cells with Aspose.Cells (.NET)
// Description: C# example that loads an XLSX workbook using CultureInfo "pt-BR", scans all worksheets, identifies cells with a percent style, and replaces the built‑in format with the localized custom format "#.##0,00%" before saving the file.
// Keywords: Aspose.Cells | C# | .NET | custom percentage format | pt-BR CultureInfo | Brazilian Portuguese number format | Style.IsPercent | Excel localization | LoadOptions.CultureInfo | custom number format string
// Common Searches: Aspose.Cells apply Brazilian percent format | C# set custom percentage format pt-BR | How to use CultureInfo with Aspose.Cells | Replace built‑in percent style in Excel using Aspose | Localized number formatting in .NET Excel library
// Developer Intent: Replace every built‑in percent style in a workbook with the Brazilian Portuguese custom format "#.##0,00%" using Aspose.Cells.
// Use Cases: Create a sample workbook, apply the default percent style, and test the conversion to a localized format. | Load an existing Excel file with LoadOptions.CultureInfo set to "pt-BR", detect cells where Style.IsPercent is true, and assign the custom format "#.##0,00%". | Save the updated workbook while preserving Brazilian Portuguese numeric conventions for downstream processing or reporting.
// AI Prompts: Generate C# code that opens an XLSX file with Aspose.Cells using pt-BR CultureInfo and changes all percent‑styled cells to the custom format "#.##0,00%". | Explain the purpose of Style.IsPercent in Aspose.Cells and show how to apply a locale‑specific custom number format for percentages. | Provide a step‑by‑step guide to localize number formats in Excel workbooks with Aspose.Cells for Brazilian Portuguese.

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // C# example that loads an XLSX workbook using CultureInfo "pt-BR", scans all worksheets, identifies cells with a percent style, and replaces the built‑in format with the localized custom format "#.##0,00%" before saving the file.
    public class ApplyCustomPercentageFormat
    {
        public static void Main()
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

            // Ensure the input file exists; create a sample workbook if missing
            if (!File.Exists(inputPath))
            {
                var sampleWb = new Workbook();
                var sheet = sampleWb.Worksheets[0];
                var cell = sheet.Cells["A1"];
                cell.PutValue(0.1234); // 12.34%

                // Apply a built‑in percent format (index 10) to the sample cell
                var style = cell.GetStyle();
                style.Number = 10; // Built‑in percent format
                cell.SetStyle(style);

                sampleWb.Save(inputPath, SaveFormat.Xlsx);
            }

            // Load the workbook with Brazilian Portuguese culture
            var loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                CultureInfo = new CultureInfo("pt-BR")
            };
            var workbook = new Workbook(inputPath, loadOptions);

            // Custom percentage format for pt-BR (decimal separator ',' and group separator '.')
            const string customPercentageFormat = "#.##0,00%";

            // Apply the custom format to all percentage cells
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];
                        if (cell == null || cell.Type == CellValueType.IsNull)
                            continue;

                        Style style = cell.GetStyle();

                        // Check if the cell already uses a percent format
                        if (style.IsPercent)
                        {
                            // Apply the custom percentage format
                            style.Custom = customPercentageFormat;
                            cell.SetStyle(style);
                        }
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}
