// Title: Export a Named Range to CSV with Aspose.Cells for .NET (C#)
// Description: Load an Excel workbook, retrieve the "EmployeeList" named range using GetRangeByName, convert it to a DataTable with ExportDataTable, and write the data to a CSV file. The example handles header creation, proper escaping of commas and quotes, and includes error handling for missing files or ranges.
// Keywords: Aspose.Cells export named range | C# export named range to CSV | GetRangeByName Aspose.Cells | ExportDataTable to CSV | Excel named range CSV conversion | .NET write DataTable as CSV | EmployeeList named range export | Aspose.Cells CSV output | flat file export from Excel | C# Excel to CSV example
// Common Searches: how to export a named range from Excel to CSV using Aspose.Cells C# | Aspose.Cells GetRangeByName example for CSV | C# convert Excel named range to CSV file | export EmployeeList range to CSV with Aspose.Cells | write DataTable to CSV in .NET
// Developer Intent: Generate a CSV file that contains only the data from the "EmployeeList" named range in an Excel workbook.
// Use Cases: Produce a lightweight CSV report of employee records for HR analytics. | Supply a flat‑file feed to an external system that requires only the EmployeeList data. | Create a backup of a specific worksheet segment for data archiving or migration.
// AI Prompts: Write C# code using Aspose.Cells to export any named range to a CSV file with a custom delimiter and UTF‑8 encoding. | Modify the example to add a UTF‑8 BOM, handle null values, and allow the delimiter to be passed as a parameter. | Create a reusable method that takes workbook path, named range name, output CSV path, and returns a success flag with detailed logging.

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace ExportNamedRangeToCsv
{
    // Load an Excel workbook, retrieve the "EmployeeList" named range using GetRangeByName, convert it to a DataTable with ExportDataTable, and write the data to a CSV file. The example handles header creation, proper escaping of commas and quotes, and includes error handling for missing files or ranges.
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file containing the named range "EmployeeList"
            string excelPath = "EmployeeData.xlsx";

            // Path for the resulting CSV file
            string csvPath = "EmployeeList.csv";

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

                // Retrieve the named range "EmployeeList"
                // GetRangeByName returns an Aspose.Cells.Range object.
                Aspose.Cells.Range employeeRange = workbook.Worksheets.GetRangeByName("EmployeeList");
                if (employeeRange == null)
                {
                    Console.WriteLine("Error: Named range \"EmployeeList\" was not found in the workbook.");
                    return;
                }

                // Export the range data to a DataTable
                DataTable dataTable = employeeRange.ExportDataTable();

                // Write the DataTable contents to a CSV file
                using (StreamWriter writer = new StreamWriter(csvPath))
                {
                    // Write header row
                    for (int col = 0; col < dataTable.Columns.Count; col++)
                    {
                        writer.Write(dataTable.Columns[col].ColumnName);
                        if (col < dataTable.Columns.Count - 1)
                            writer.Write(",");
                    }
                    writer.WriteLine();

                    // Write data rows
                    foreach (DataRow row in dataTable.Rows)
                    {
                        for (int col = 0; col < dataTable.Columns.Count; col++)
                        {
                            // Escape commas and quotes in field values
                            string field = row[col]?.ToString() ?? string.Empty;
                            if (field.Contains(",") || field.Contains("\""))
                            {
                                field = $"\"{field.Replace("\"", "\"\"")}\"";
                            }
                            writer.Write(field);
                            if (col < dataTable.Columns.Count - 1)
                                writer.Write(",");
                        }
                        writer.WriteLine();
                    }
                }

                Console.WriteLine($"Named range \"EmployeeList\" has been exported to CSV at: {csvPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
