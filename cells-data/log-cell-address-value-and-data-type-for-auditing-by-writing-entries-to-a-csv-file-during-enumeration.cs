// Title: Create a CSV audit log of cell addresses, values, and data types using Aspose.Cells for .NET
// AI Prompts: Write a C# console program that opens an .xlsx file with Aspose.Cells, loops through every used cell, and appends a line containing the cell's address, stringified value, and Aspose.Cells Type enum to a CSV file, ensuring proper CSV escaping. | Generate code that extracts the Name, Value, and Type of each cell in all worksheets of a workbook and saves the information to a comma‑separated audit file, handling commas, quotes, and line breaks in cell values. | Provide a method that iterates a workbook's Cells collection, formats each entry as "Address,Value,DataType", and writes it to a CSV stream using StreamWriter with correct escaping for special characters.
// Common Searches: aspnet how to write Excel cell address and type to CSV using Aspose.Cells | c# export all cell values with data types from workbook to audit csv | aspose.cells enumerate cells and create audit log file | csv escaping commas quotes line breaks when exporting Excel data in C# | log worksheet cell information to CSV for compliance auditing Aspose.Cells
// Tags: Aspose.Cells export cells to CSV | enumerate workbook cells Aspose.Cells | CSV audit logging .NET | escape special characters CSV C# | cell data type extraction Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an .xlsx workbook with Aspose.Cells, iterates every used cell across all worksheets, captures each cell's address, value, and data type, applies CSV‑compliant escaping, and writes the records to a comma‑separated audit file.
class CellAuditExporter
{
    static void Main()
    {
        // Path to the source Excel file
        string excelPath = "input.xlsx";

        // Path to the output CSV file
        string csvPath = "audit_log.csv";

        // Load the workbook using Aspose.Cells
        Workbook workbook = new Workbook(excelPath);

        // Create a StreamWriter for the CSV output
        using (StreamWriter writer = new StreamWriter(csvPath, false))
        {
            // Write CSV header
            writer.WriteLine("Address,Value,DataType");

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through each cell in the used range of the worksheet
                foreach (Cell cell in sheet.Cells)
                {
                    // Get cell address (e.g., "A1")
                    string address = cell.Name;

                    // Get cell value as string (handle nulls)
                    string value = cell.Value?.ToString() ?? string.Empty;

                    // Get cell data type (enum name)
                    string dataType = cell.Type.ToString();

                    // Escape commas and quotes in the value for CSV compliance
                    if (value.Contains("\""))
                        value = value.Replace("\"", "\"\"");

                    if (value.Contains(",") || value.Contains("\"") || value.Contains("\n"))
                        value = $"\"{value}\"";

                    // Write the CSV line
                    writer.WriteLine($"{address},{value},{dataType}");
                }
            }
        }

        Console.WriteLine($"Audit log has been written to '{csvPath}'.");
    }
}
