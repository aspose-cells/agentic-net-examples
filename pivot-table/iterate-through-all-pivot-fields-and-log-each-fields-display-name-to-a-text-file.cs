// Title: Log every pivot field’s display name from an Excel workbook to a text file using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells iterate all pivot fields and write their display names to a txt file | How to export pivot table field names to a text file with Aspose.Cells in .NET | Save Excel pivot field display names to a log file using Aspose.Cells library | Aspose.Cells code to list row, column, page, and data pivot fields in C# | Write pivot field DisplayName values to a file with Aspose.Cells for .NET
// Tags: aspocells export pivot field display names | c# iterate pivot tables aspocells | write pivot field names to txt file | log excel pivot fields using Aspose.Cells | pivotfieldcollection enumeration in .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // Required for PivotTable and PivotField classes

// The example loads an Excel workbook with Aspose.Cells, walks through each worksheet and its pivot tables, and writes the DisplayName of every row, column, page, and data pivot field to a specified text file, creating the output folder if needed and handling common I/O errors.
class Program
{
    static void Main()
    {
        // Path to the Excel file containing pivot tables
        string excelPath = @"C:\Input\Sample.xlsx";

        // Path to the output text file where field names will be logged
        string logPath = @"C:\Output\PivotFields.txt";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(excelPath))
                throw new FileNotFoundException($"Input file not found: {excelPath}");

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(logPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Load the workbook
            Workbook workbook = new Workbook(excelPath);

            // Prepare a StreamWriter for logging
            using (StreamWriter writer = new StreamWriter(logPath, false))
            {
                // Iterate through each worksheet in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through each pivot table on the worksheet
                    foreach (PivotTable pivotTable in sheet.PivotTables)
                    {
                        // Helper local function to write field names from a collection
                        void WriteFields(PivotFieldCollection fields)
                        {
                            if (fields == null) return;
                            foreach (PivotField field in fields)
                            {
                                // Write the display name of the field to the log file
                                writer.WriteLine(field.DisplayName);
                            }
                        }

                        // Write fields from all possible collections
                        WriteFields(pivotTable.RowFields);
                        WriteFields(pivotTable.ColumnFields);
                        WriteFields(pivotTable.PageFields);
                        WriteFields(pivotTable.DataFields);
                    }
                }
            }

            Console.WriteLine("Pivot field display names have been logged to: " + logPath);
        }
        catch (Exception ex)
        {
            // Log or display the error details
            Console.Error.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
