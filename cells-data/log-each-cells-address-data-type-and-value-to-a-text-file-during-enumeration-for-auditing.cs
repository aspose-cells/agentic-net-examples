// Title: Generate a tab‑delimited audit log of cell addresses, data types, and values while iterating worksheets with Aspose.Cells for .NET
// Common Searches: aspnet c# write cell address datatype and value to txt using Aspose.Cells | enumerate used range of an Excel worksheet and export cell details with Aspose.Cells .NET | audit Excel workbook cells and create log file when workbook is password protected Aspose.Cells | how to skip null cells and log only populated cells with Aspose.Cells in C# | save cell audit information to a text file from multiple worksheets using Aspose.Cells
// Tags: generate cell audit file with Aspose.Cells | C# iterate worksheets and log cell metadata | Aspose.Cells open encrypted workbook | skip null cells during Aspose.Cells enumeration | tab delimited cell log Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsAudit
{
    // Loads an Excel workbook (prompting for a password if required), iterates each worksheet's used range, skips null cells, and writes the cell address, data type, and value as a tab‑delimited line to a text file, creating a comprehensive audit log.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file to be audited
            string excelPath = "input.xlsx";

            // Path to the output text file where audit information will be written
            string logPath = "CellAuditLog.txt";

            // Verify that the input file exists
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Input file not found: {excelPath}");
                return;
            }

            Workbook workbook = null;

            try
            {
                // Try to load the workbook without a password
                workbook = new Workbook(excelPath);
            }
            catch (CellsException ex) when (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                // Workbook is password protected; prompt for password and reload
                Console.WriteLine("The workbook is password protected.");
                Console.Write("Enter password: ");
                string password = Console.ReadLine() ?? string.Empty;

                if (string.IsNullOrEmpty(password))
                {
                    Console.WriteLine("No password entered. Cannot open the workbook.");
                    return;
                }

                try
                {
                    var loadOptions = new LoadOptions(LoadFormat.Xlsx) { Password = password };
                    workbook = new Workbook(excelPath, loadOptions);
                }
                catch (Exception loadEx)
                {
                    Console.WriteLine($"Failed to load workbook with provided password: {loadEx.Message}");
                    return;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading workbook: {ex.Message}");
                return;
            }

            try
            {
                // Prepare a StreamWriter for logging
                using (StreamWriter writer = new StreamWriter(logPath, false))
                {
                    // Iterate through each worksheet in the workbook
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        writer.WriteLine($"Worksheet: {sheet.Name}");
                        writer.WriteLine("Address\tDataType\tValue");

                        // Determine the used range to avoid iterating over the entire sheet
                        int maxRow = sheet.Cells.MaxDataRow;
                        int maxCol = sheet.Cells.MaxDataColumn;

                        // Enumerate cells within the used range
                        for (int row = 0; row <= maxRow; row++)
                        {
                            for (int col = 0; col <= maxCol; col++)
                            {
                                Cell cell = sheet.Cells[row, col];

                                // Skip cells that have never been used (type is Null)
                                if (cell.Type == CellValueType.IsNull)
                                    continue;

                                string address = cell.Name;                     // e.g., "A1"
                                string dataType = cell.Type.ToString();         // e.g., "IsNumeric"
                                string value = cell.Value?.ToString() ?? string.Empty;

                                writer.WriteLine($"{address}\t{dataType}\t{value}");
                            }
                        }

                        // Add a blank line after each worksheet for readability
                        writer.WriteLine();
                    }
                }

                Console.WriteLine($"Cell audit completed. Log saved to: {logPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during audit processing: {ex.Message}");
            }
        }
    }
}
