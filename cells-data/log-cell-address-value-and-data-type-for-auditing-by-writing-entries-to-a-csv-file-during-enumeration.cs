// Title: How to enumerate cells in an Excel worksheet and write their address, value, and data type to a CSV file using Aspose.Cells for .NET (C#)
// Common Searches: C# loop through all used cells in an Excel file and save address, value and type to a CSV with Aspose.Cells | Aspose.Cells export cell address and data type to CSV for audit purposes in .NET | Write Excel cell metadata (address, value, datatype) to a CSV file using Aspose.Cells library | Handle commas and line breaks when exporting Excel cell values to CSV with Aspose.Cells | Detect empty worksheet while logging cell details to CSV using Aspose.Cells C#
// Tags: Aspose.Cells enumerate used range cells | write cell address and value to CSV C# | export cell data type to CSV Aspose.Cells | CSV escaping for commas and quotes Aspose.Cells | audit Excel worksheet cells with Aspose.Cells | load first worksheet from workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, checks for a first worksheet, obtains its used range, and iterates each cell. For every cell it records the address, stringified value, and the cell's data type, applying proper CSV escaping for commas, quotes, and line breaks. The collected information is written to a CSV file with a header, handling empty sheets and missing files gracefully.
class CellAudit
{
    static void Main()
    {
        // Paths to the source Excel file and the output CSV file
        string excelPath = @"C:\Data\SourceWorkbook.xlsx";
        string csvPath   = @"C:\Data\CellAudit.csv";

        // Verify that the source Excel file exists
        if (!File.Exists(excelPath))
        {
            Console.WriteLine($"Error: The file \"{excelPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(excelPath);

            // Ensure the workbook contains at least one worksheet
            if (workbook.Worksheets.Count == 0)
            {
                Console.WriteLine("Error: The workbook does not contain any worksheets.");
                return;
            }

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Get the used range of the worksheet; may be null if the sheet is empty
            var usedRange = sheet.Cells.MaxDisplayRange; // 'var' avoids ambiguity with System.Range

            // Prepare the CSV file for writing
            using (StreamWriter writer = new StreamWriter(csvPath, false))
            {
                // Write CSV header
                writer.WriteLine("Address,Value,DataType");

                // If there is no used range, the sheet is empty – write only the header
                if (usedRange == null)
                {
                    Console.WriteLine("Info: The worksheet is empty. Only header was written to the CSV file.");
                    return;
                }

                // Iterate through each cell in the used range
                int startRow = usedRange.FirstRow;
                int startCol = usedRange.FirstColumn;
                int rowCount = usedRange.RowCount;
                int colCount = usedRange.ColumnCount;

                for (int r = 0; r < rowCount; r++)
                {
                    for (int c = 0; c < colCount; c++)
                    {
                        Cell cell = sheet.Cells[startRow + r, startCol + c];

                        // Cell address (e.g., A1)
                        string address = cell.Name;

                        // Cell value as string (handle nulls)
                        string value = cell.Value?.ToString() ?? string.Empty;

                        // Escape double quotes and commas in the value
                        if (value.Contains("\""))
                            value = value.Replace("\"", "\"\"");
                        if (value.Contains(",") || value.Contains("\n") || value.Contains("\r"))
                            value = $"\"{value}\"";

                        // Cell data type (e.g., String, Numeric, Boolean, DateTime, IsError, IsNull)
                        string dataType = cell.Type.ToString();

                        // Write the CSV line
                        writer.WriteLine($"{address},{value},{dataType}");
                    }
                }
            }

            Console.WriteLine($"Audit completed successfully. CSV saved to \"{csvPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
