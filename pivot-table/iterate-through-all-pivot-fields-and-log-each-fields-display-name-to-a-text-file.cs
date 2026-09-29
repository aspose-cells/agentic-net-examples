// Title: Log each pivot table field’s display name to a text file using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an Excel workbook with Aspose.Cells, iterates all pivot tables, and writes every field’s DisplayName to a .txt log file. | Extend the sample to also capture hidden pivot fields and output the names in CSV format instead of plain text. | Create a reusable method that returns a List<string> of display names for all row, column, page, and data fields of a given PivotTable using Aspose.Cells.
// Common Searches: Aspose.Cells C# export pivot table field names to a text file | How to enumerate all fields of a pivot table in a .NET workbook | Write pivot field display names to a log file with Aspose.Cells | Get row, column, page, and data field names from Excel pivot tables using C#
// Tags: Aspose.Cells enumerate pivot fields | write pivot field names to txt C# | log pivot table metadata .NET | extract pivot field display names Aspose | pivot table field iteration Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // Required for PivotTable and PivotField classes

// The example loads an Excel workbook with Aspose.Cells, checks the file's existence, then walks through each worksheet and its pivot tables. For every pivot table it writes the DisplayName of row, column, page, and data fields to a text file, handling load and write errors gracefully.
class PivotFieldLogger
{
    static void Main()
    {
        // Path to the input workbook
        string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Path for the output log file
        string logPath = "PivotFieldsLog.txt";

        try
        {
            // Write pivot field display names to the log file
            using (StreamWriter writer = new StreamWriter(logPath, false))
            {
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (PivotTable pivotTable in sheet.PivotTables)
                    {
                        LogPivotFields(pivotTable, writer);
                    }
                }
            }

            Console.WriteLine("Pivot field display names have been logged to: " + Path.GetFullPath(logPath));
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred while writing the log file: {ex.Message}");
        }
    }

    // Logs display names of all fields used in a pivot table
    private static void LogPivotFields(PivotTable pivotTable, StreamWriter writer)
    {
        try
        {
            foreach (PivotField field in pivotTable.RowFields)
                writer.WriteLine(field.DisplayName);

            foreach (PivotField field in pivotTable.ColumnFields)
                writer.WriteLine(field.DisplayName);

            foreach (PivotField field in pivotTable.PageFields)
                writer.WriteLine(field.DisplayName);

            foreach (PivotField field in pivotTable.DataFields)
                writer.WriteLine(field.DisplayName);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to log fields for a pivot table: {ex.Message}");
        }
    }
}
