// Title: Convert HTML to XLSX while preserving CSS‑based conditional formatting using Aspose.Cells for .NET (C#)
// Common Searches: how to import CSS conditional formatting from HTML to Excel with Aspose.Cells in C# | Aspose.Cells HtmlLoadOptions ImportConditionalFormatting property example | C# create conditional formatting rules based on HTML cell styles | preserve background and font colors when converting HTML tables to XLSX using Aspose.Cells | programmatically convert HTML to Excel and retain style‑based formatting
// Tags: HtmlLoadOptions.ImportConditionalFormatting | create conditional formatting from CSS styles Aspose.Cells | C# group cells by background and font color | convert HTML table to XLSX preserving styles | A1 address parsing to row column indices C#

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace HtmlToExcelWithConditionalFormatting
{
    // The sample loads an HTML file into an Aspose.Cells Workbook with HtmlLoadOptions (enabling ImportConditionalFormatting when available), scans all cells to identify non‑default background or font colors, groups cells sharing the same style, creates a conditional formatting collection for each group using a dummy TRUE expression and applies the representative style, and finally saves the workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            // Validate arguments
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: HtmlToExcelWithConditionalFormatting <inputHtmlPath> <outputExcelPath>");
                return;
            }

            string inputHtmlPath = args[0];
            string outputExcelPath = args[1];

            if (!File.Exists(inputHtmlPath))
            {
                Console.WriteLine($"Error: Input HTML file not found: {inputHtmlPath}");
                return;
            }

            try
            {
                // -----------------------------------------------------------------
                // 1. Load the HTML file into a Workbook.
                // -----------------------------------------------------------------
                HtmlLoadOptions loadOptions = new HtmlLoadOptions();

                // Enable ImportConditionalFormatting if supported (Aspose.Cells 22.12+)
                var importCondProp = typeof(HtmlLoadOptions).GetProperty("ImportConditionalFormatting");
                if (importCondProp != null && importCondProp.CanWrite)
                {
                    importCondProp.SetValue(loadOptions, true);
                    Console.WriteLine("HtmlLoadOptions.ImportConditionalFormatting set to true.");
                }
                else
                {
                    Console.WriteLine("HtmlLoadOptions.ImportConditionalFormatting not available – will handle manually.");
                }

                // Load the workbook from HTML.
                Workbook workbook = new Workbook(inputHtmlPath, loadOptions);
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // -----------------------------------------------------------------
                // 2. Collect cells that have non‑default styling (originating from CSS).
                // -----------------------------------------------------------------
                var styleGroups = new Dictionary<string, List<string>>();

                for (int row = cells.MinRow; row <= cells.MaxRow; row++)
                {
                    for (int col = cells.MinColumn; col <= cells.MaxColumn; col++)
                    {
                        Cell cell = cells[row, col];
                        if (cell == null) continue;

                        Style style = cell.GetStyle();

                        bool hasBackground = style.BackgroundColor != Color.Empty && style.Pattern != BackgroundType.None;
                        bool hasForeground = style.Font.Color != Color.Empty;

                        if (!hasBackground && !hasForeground) continue;

                        string key = $"{style.BackgroundColor.ToArgb()}|{style.Font.Color.ToArgb()}|{style.Pattern}";
                        string address = cell.Name;

                        if (!styleGroups.ContainsKey(key))
                            styleGroups[key] = new List<string>();

                        styleGroups[key].Add(address);
                    }
                }

                // -----------------------------------------------------------------
                // 3. Create conditional formatting rules for each style group.
                // -----------------------------------------------------------------
                foreach (var kvp in styleGroups)
                {
                    List<string> addresses = kvp.Value;
                    if (addresses.Count < 2) continue; // skip singletons

                    // Add a new conditional formatting collection.
                    int cfIndex = sheet.ConditionalFormattings.Add();
                    FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];

                    // Add each cell address as a separate area.
                    foreach (string addr in addresses)
                    {
                        int row = CellPositionToRow(addr);
                        int col = CellPositionToColumn(addr);
                        CellArea area = new CellArea
                        {
                            StartRow = row,
                            StartColumn = col,
                            EndRow = row,
                            EndColumn = col
                        };
                        fcc.AddArea(area);
                    }

                    // Dummy condition that always evaluates to TRUE.
                    int conditionIndex = fcc.AddCondition(FormatConditionType.Expression, OperatorType.None, "TRUE", null);

                    // Use the style from the first cell as the representative style.
                    Cell sampleCell = cells[CellPositionToRow(addresses[0]), CellPositionToColumn(addresses[0])];
                    Style sampleStyle = sampleCell.GetStyle();

                    // Apply the style to the condition.
                    fcc[conditionIndex].Style = sampleStyle;
                }

                // -----------------------------------------------------------------
                // 4. Save the workbook as an Excel file.
                // -----------------------------------------------------------------
                workbook.Save(outputExcelPath, SaveFormat.Xlsx);
                Console.WriteLine($"HTML successfully converted to Excel: {outputExcelPath}");
                Console.WriteLine("Conditional formatting rules derived from CSS have been applied.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during conversion:");
                Console.WriteLine(ex.Message);
            }
        }

        // Helper to convert an A1 style address (e.g., "B12") to a zero‑based row index.
        private static int CellPositionToRow(string address)
        {
            string clean = address.Replace("$", "");
            int i = 0;
            while (i < clean.Length && Char.IsLetter(clean[i])) i++;
            string rowPart = clean.Substring(i);
            return int.Parse(rowPart) - 1; // zero‑based
        }

        // Helper to convert an A1 style address (e.g., "B12") to a zero‑based column index.
        private static int CellPositionToColumn(string address)
        {
            string clean = address.Replace("$", "");
            int i = 0;
            while (i < clean.Length && Char.IsLetter(clean[i])) i++;
            string colPart = clean.Substring(0, i).ToUpperInvariant();

            int columnNumber = 0;
            foreach (char c in colPart)
            {
                columnNumber = columnNumber * 26 + (c - 'A' + 1);
            }
            return columnNumber - 1; // zero‑based
        }
    }
}
