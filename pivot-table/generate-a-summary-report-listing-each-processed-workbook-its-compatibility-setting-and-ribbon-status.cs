// Title: Generate a C# summary workbook that lists each Excel file’s Excel2003 compatibility mode and ribbon enabled flag using Aspose.Cells
// Common Searches: C# Aspose.Cells read IsExcel2003CompatibilityMode property from a workbook | How to retrieve ribbon enabled status for multiple Excel files with Aspose.Cells | Create a batch summary report of workbook settings (compatibility and ribbon) in .NET | Aspose.Cells reflection example to access workbook Settings properties in C# | Generate an Excel summary sheet of file paths and workbook flags using Aspose.Cells
// Tags: Aspose.Cells read workbook compatibility mode | Aspose.Cells retrieve ribbon enabled flag | C# batch workbook settings report | Aspose.Cells generate summary worksheet | reflection access workbook Settings properties

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace WorkbookSummaryReport
{
    // The C# program iterates over a list of Excel file paths, loads each workbook with Aspose.Cells, and uses reflection to read the IsExcel2003CompatibilityMode and IsRibbonEnabled properties from the workbook's Settings. It records the file path, compatibility flag, and ribbon status (or error indicators) into a new summary workbook, auto‑fits the columns, and saves the report, handling missing files and exceptions gracefully.
    class Program
    {
        static void Main(string[] args)
        {
            // List of workbook file paths to process
            List<string> workbookFiles = new List<string>
            {
                @"C:\Workbooks\Book1.xlsx",
                @"C:\Workbooks\Book2.xlsx",
                @"C:\Workbooks\Book3.xlsx"
                // Add more file paths as needed
            };

            // Create a new workbook for the summary report
            Workbook summaryWorkbook = new Workbook();

            // Access the first worksheet where the report will be written
            Worksheet sheet = summaryWorkbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Write header row
            cells["A1"].PutValue("Workbook Path");
            cells["B1"].PutValue("Compatibility (Excel2003)");
            cells["C1"].PutValue("Ribbon Enabled");

            int currentRow = 1; // zero‑based index; row 1 is the second row (after header)

            foreach (string filePath in workbookFiles)
            {
                try
                {
                    // Verify that the file exists before attempting to load it
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found: {filePath}");
                        cells[currentRow, 0].PutValue(filePath);
                        cells[currentRow, 1].PutValue("File not found");
                        cells[currentRow, 2].PutValue("File not found");
                        currentRow++;
                        continue;
                    }

                    // Load each workbook using the constructor that accepts a file path
                    Workbook wb = new Workbook(filePath);

                    // Retrieve compatibility setting (Excel2003 compatibility mode) via reflection
                    bool isExcel2003Compat = false;
                    try
                    {
                        var compatProp = wb.Settings.GetType().GetProperty("IsExcel2003CompatibilityMode");
                        if (compatProp != null && compatProp.PropertyType == typeof(bool))
                        {
                            object value = compatProp.GetValue(wb.Settings);
                            if (value != null)
                                isExcel2003Compat = (bool)value;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Unable to read Excel2003 compatibility for '{filePath}': {ex.Message}");
                    }

                    // Retrieve ribbon status – not available in older versions, default to false
                    bool isRibbonEnabled = false;
                    try
                    {
                        var ribbonProp = wb.Settings.GetType().GetProperty("IsRibbonEnabled");
                        if (ribbonProp != null && ribbonProp.PropertyType == typeof(bool))
                        {
                            object value = ribbonProp.GetValue(wb.Settings);
                            if (value != null)
                                isRibbonEnabled = (bool)value;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Unable to read Ribbon status for '{filePath}': {ex.Message}");
                    }

                    // Write data to the summary sheet
                    cells[currentRow, 0].PutValue(filePath);
                    cells[currentRow, 1].PutValue(isExcel2003Compat);
                    cells[currentRow, 2].PutValue(isRibbonEnabled);
                }
                catch (Exception ex)
                {
                    // Log the exception and continue processing other files
                    Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
                    cells[currentRow, 0].PutValue(filePath);
                    cells[currentRow, 1].PutValue("Error");
                    cells[currentRow, 2].PutValue("Error");
                }

                currentRow++;
            }

            // Auto‑fit columns for better readability
            sheet.AutoFitColumns();

            // Save the summary report
            try
            {
                string summaryPath = @"C:\Workbooks\WorkbookSummaryReport.xlsx";
                string summaryDir = Path.GetDirectoryName(summaryPath);
                if (!string.IsNullOrEmpty(summaryDir) && !Directory.Exists(summaryDir))
                {
                    Directory.CreateDirectory(summaryDir);
                }

                summaryWorkbook.Save(summaryPath);
                Console.WriteLine($"Summary report saved to: {summaryPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save summary report: {ex.Message}");
            }
        }
    }
}
