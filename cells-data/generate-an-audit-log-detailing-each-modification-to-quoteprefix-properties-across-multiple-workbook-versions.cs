// Title: C# program to audit QuotePrefix changes across multiple Excel workbook versions using Aspose.Cells
// Common Searches: how to detect QuotePrefix flag changes in Excel workbooks with Aspose.Cells C# | log leading apostrophe modifications across .xlsx file versions using Aspose.Cells | compare cell QuotePrefix property between different workbook revisions in .NET | generate audit trail for cell QuotePrefix changes across multiple Excel files C# | track QuotePrefix flag updates over time with Aspose.Cells API
// Tags: Aspose.Cells detect QuotePrefix flag | C# audit cell leading apostrophe | compare Excel workbook versions QuotePrefix | log cell property changes Aspose.Cells | track QuotePrefix across .xlsx files

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The C# example loads a chronological list of .xlsx files with Aspose.Cells, iterates every worksheet and cell, determines each cell's QuotePrefix flag via the IsQuotePrefix property (or a fallback apostrophe check), compares it to the previously recorded state, and writes a text‑based audit log that records the initial QuotePrefix value and any subsequent changes while handling missing files and load errors.
class QuotePrefixAudit
{
    static void Main()
    {
        // List workbook versions in chronological order
        var workbookFiles = new List<string>
        {
            "Version1.xlsx",
            "Version2.xlsx",
            "Version3.xlsx"
        };

        // Store the previous QuotePrefix state for each cell (key = SheetName!CellName)
        var previousState = new Dictionary<string, bool>();

        // Create audit log file
        using (var logWriter = new StreamWriter("QuotePrefixAuditLog.txt"))
        {
            foreach (var filePath in workbookFiles)
            {
                // Verify that the workbook file exists
                if (!File.Exists(filePath))
                {
                    logWriter.WriteLine($"{DateTime.UtcNow:u} | Error: File not found - {filePath}");
                    continue;
                }

                Workbook workbook;
                try
                {
                    // Load workbook (using default load options)
                    workbook = new Workbook(filePath);
                }
                catch (Exception ex)
                {
                    logWriter.WriteLine($"{DateTime.UtcNow:u} | Error loading workbook {Path.GetFileName(filePath)}: {ex.Message}");
                    continue;
                }

                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    var cells = sheet.Cells;
                    int maxRow = cells.MaxDataRow;
                    int maxColumn = cells.MaxDataColumn;

                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxColumn; col++)
                        {
                            var cell = cells[row, col];

                            // Determine QuotePrefix flag.
                            // Aspose.Cells newer versions expose Cell.IsQuotePrefix.
                            // For compatibility, fall back to checking a leading apostrophe in the string value.
                            bool currentQuotePrefix;
                            try
                            {
                                // Attempt to use the native property if available (via reflection to avoid compile‑time dependency)
                                var prop = typeof(Cell).GetProperty("IsQuotePrefix");
                                if (prop != null && prop.PropertyType == typeof(bool))
                                {
                                    var value = prop.GetValue(cell);
                                    currentQuotePrefix = value != null && (bool)value;
                                }
                                else
                                {
                                    // Fallback detection using the cell's string representation
                                    var str = cell.StringValue;
                                    currentQuotePrefix = !string.IsNullOrEmpty(str) && str.StartsWith("'");
                                }
                            }
                            catch
                            {
                                // In case of unexpected errors, assume false
                                currentQuotePrefix = false;
                            }

                            string cellKey = $"{sheet.Name}!{cell.Name}";

                            if (previousState.TryGetValue(cellKey, out bool previousQuotePrefix))
                            {
                                if (previousQuotePrefix != currentQuotePrefix)
                                {
                                    // Log modification
                                    logWriter.WriteLine($"{DateTime.UtcNow:u} | Workbook: {Path.GetFileName(filePath)} | Cell: {cellKey} | QuotePrefix changed from {previousQuotePrefix} to {currentQuotePrefix}");
                                }
                            }
                            else
                            {
                                // First occurrence of this cell, log its initial state
                                logWriter.WriteLine($"{DateTime.UtcNow:u} | Workbook: {Path.GetFileName(filePath)} | Cell: {cellKey} | Initial QuotePrefix = {currentQuotePrefix}");
                            }

                            // Update state for next comparison
                            previousState[cellKey] = currentQuotePrefix;
                        }
                    }
                }
            }
        }
    }
}
