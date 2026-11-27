// Title: Programmatically update the data source of all Waterfall charts after bulk CSV import using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells replace waterfall chart series after importing CSV data | how to refresh multiple waterfall charts in an Excel workbook with Aspose.Cells | import bulk CSV to a sheet and reassign chart ranges in .NET | set NSeries range for waterfall charts programmatically using Aspose.Cells | automate chart data source change after bulk data load in C#
// Tags: waterfall chart data source update Aspose.Cells | bulk CSV import to worksheet Aspose.Cells | programmatic NSeries range assignment C# | iterate and modify workbook charts Aspose.Cells | clear and add chart series Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace UpdateWaterfallCharts
{
    // The example loads an existing workbook, creates or clears a "Data" sheet, imports bulk CSV data into it, determines the used range, and then iterates through every worksheet to locate Waterfall charts. For each chart it clears existing series, adds a new NSeries with the computed range, sets the category (X‑axis) data to the first column, and saves the updated workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Expect three arguments: input workbook path, bulk data CSV path, output workbook path
            if (args.Length != 3)
            {
                Console.WriteLine("Usage: UpdateWaterfallCharts <inputWorkbook.xlsx> <bulkData.csv> <outputWorkbook.xlsx>");
                return;
            }

            string inputWorkbookPath = args[0];
            string bulkDataCsvPath = args[1];
            string outputWorkbookPath = args[2];

            if (!File.Exists(inputWorkbookPath))
            {
                Console.WriteLine($"Error: Input workbook not found at '{inputWorkbookPath}'.");
                return;
            }

            if (!File.Exists(bulkDataCsvPath))
            {
                Console.WriteLine($"Error: Bulk data CSV not found at '{bulkDataCsvPath}'.");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputWorkbookPath);

                // Perform bulk data import into a dedicated sheet named "Data"
                Worksheet dataSheet = GetOrCreateDataSheet(workbook);
                ImportCsvToWorksheet(bulkDataCsvPath, dataSheet);

                // Determine the data range that will be used for the waterfall charts
                string dataRange = GetDataRangeAddress(dataSheet);

                // Update all Waterfall charts in the workbook to use the new data range
                UpdateAllWaterfallCharts(workbook, dataRange);

                // Save the modified workbook
                workbook.Save(outputWorkbookPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputWorkbookPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An unexpected error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }

        private static Worksheet GetOrCreateDataSheet(Workbook workbook)
        {
            const string dataSheetName = "Data";
            foreach (Worksheet ws in workbook.Worksheets)
            {
                if (ws.Name.Equals(dataSheetName, StringComparison.OrdinalIgnoreCase))
                {
                    // Clear existing content if any
                    int maxRow = ws.Cells.MaxDataRow;
                    int maxCol = ws.Cells.MaxDataColumn;
                    if (maxRow >= 0 && maxCol >= 0)
                    {
                        ws.Cells.ClearContents(0, 0, maxRow, maxCol);
                    }
                    return ws;
                }
            }

            // Sheet not found – create a new one at the end
            int index = workbook.Worksheets.Add();
            Worksheet newSheet = workbook.Worksheets[index];
            newSheet.Name = dataSheetName;
            return newSheet;
        }

        private static void ImportCsvToWorksheet(string csvPath, Worksheet sheet)
        {
            try
            {
                // Load CSV workbook
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
                Workbook csvWorkbook = new Workbook(csvPath, loadOptions);

                // Copy rows from CSV worksheet to target sheet
                Cells sourceCells = csvWorkbook.Worksheets[0].Cells;
                int rowCount = sourceCells.MaxDataRow + 1; // include header row
                if (rowCount > 0)
                {
                    sheet.Cells.CopyRows(sourceCells, 0, 0, rowCount);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error importing CSV: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Constructs an absolute address string (e.g., =Data!$A$2:$D$15) that covers all used cells
        /// in the provided worksheet, excluding the header row (assumed to be row 1).
        /// </summary>
        private static string GetDataRangeAddress(Worksheet sheet)
        {
            int lastRow = sheet.Cells.MaxDataRow;      // zero‑based
            int lastCol = sheet.Cells.MaxDataColumn;   // zero‑based

            if (lastRow < 1)
            {
                throw new InvalidOperationException("The data sheet does not contain enough rows for a chart (minimum 2 rows required).");
            }

            string startCell = $"${ColumnIndexToName(0)}$2"; // column A, row 2
            string endCell = $"${ColumnIndexToName(lastCol)}${lastRow + 1}"; // +1 because rows are zero‑based

            // Return a formula‑style range that can be used directly in chart methods.
            return $"=Data!{startCell}:{endCell}";
        }

        /// <summary>
        /// Converts a zero‑based column index to its Excel column letter(s).
        /// </summary>
        private static string ColumnIndexToName(int columnIndex)
        {
            const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            string result = string.Empty;
            int dividend = columnIndex + 1;

            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                result = letters[modulo] + result;
                dividend = (dividend - modulo) / 26;
            }

            return result;
        }

        /// <summary>
        /// Finds every Waterfall chart in the workbook and re‑assigns its data source to the supplied range.
        /// </summary>
        private static void UpdateAllWaterfallCharts(Workbook workbook, string dataRange)
        {
            foreach (Worksheet ws in workbook.Worksheets)
            {
                foreach (Chart chart in ws.Charts)
                {
                    if (chart.Type == ChartType.Waterfall)
                    {
                        Console.WriteLine($"Updating Waterfall chart on sheet '{ws.Name}'.");

                        // Clear any existing series to avoid duplication.
                        chart.NSeries.Clear();

                        // Remove leading '=' because NSeries.Add expects a range without it.
                        string cleanRange = dataRange.TrimStart('=');
                        chart.NSeries.Add(cleanRange, true);

                        // Set category (X‑axis) data if the first column contains labels.
                        string categoryRange = GetCategoryRangeFromDataRange(dataRange);
                        if (!string.IsNullOrEmpty(categoryRange))
                        {
                            chart.NSeries.CategoryData = categoryRange;
                        }

                        // No explicit recalculation method is required; Aspose.Cells updates the chart on save.
                    }
                }
            }
        }

        /// <summary>
        /// Derives a category range (first column of the data range) from a full data range string.
        /// Expected format: =Sheet!$A$2:$D$15
        /// Returns: =Sheet!$A$2:$A$15
        /// </summary>
        private static string GetCategoryRangeFromDataRange(string fullRange)
        {
            if (string.IsNullOrWhiteSpace(fullRange))
                return null;

            // Remove leading '=' for easier parsing.
            string trimmed = fullRange.TrimStart('=');
            int exclPos = trimmed.IndexOf('!');
            if (exclPos < 0) return null;

            string sheetPart = trimmed.Substring(0, exclPos);
            string rangePart = trimmed.Substring(exclPos + 1);

            // Split start and end addresses.
            string[] parts = rangePart.Split(':');
            if (parts.Length != 2) return null;

            string startAddr = parts[0];
            string endAddr = parts[1];

            // Extract column letters from the start address.
            int startColEnd = 0;
            while (startColEnd < startAddr.Length && !char.IsDigit(startAddr[startColEnd]))
                startColEnd++;

            if (startColEnd == 0) return null;

            string columnLetter = startAddr.Substring(0, startColEnd);
            string startRow = startAddr.Substring(startColEnd);
            string endRow = endAddr.Substring(endAddr.IndexOfAny("0123456789".ToCharArray()));

            // Build the category range.
            return $"={sheetPart}!${columnLetter}${startRow}:${columnLetter}${endRow}";
        }
    }
}
