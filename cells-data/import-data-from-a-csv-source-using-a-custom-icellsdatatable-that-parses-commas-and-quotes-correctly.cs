// Title: Import CSV with commas and quoted fields into Aspose.Cells using ICellsDataTable (C#)
// Description: Demonstrates a lightweight C# CSV parser that correctly handles commas, double quotes, and escaped quotes, builds an ICellsDataTable from the parsed rows (using the first row as a header), imports the table into a worksheet with numeric conversion and visible field names, and saves the result as an XLSX file.
// Keywords: Aspose.Cells | CSV import | quoted commas | ICellsDataTable | ImportTableOptions | ConvertNumericData | custom CSV parser | C# Excel export | .NET | Excel workbook creation
// Common Searches: How to import CSV with quoted commas into Aspose.Cells | Aspose.Cells ICellsDataTable from CSV string | ImportTableOptions ConvertNumericData example | C# parse CSV with escaped quotes for Excel | Load CSV data directly into Aspose.Cells worksheet
// Developer Intent: Parse CSV text that contains commas and quoted fields, create an ICellsDataTable from the parsed rows, and import it into an Aspose.Cells worksheet with proper data types and column headers.
// Use Cases: Convert API‑returned CSV strings into Excel files while preserving embedded commas and quotes. | Generate reports from log files stored as CSV without writing intermediate files. | Create Excel worksheets from user‑uploaded CSV data, automatically converting numeric strings and displaying column names.
// AI Prompts: Write a C# method that reads a CSV file with escaped double quotes and returns an ICellsDataTable ready for Aspose.Cells import. | Show how to set ImportTableOptions to display field names and convert numeric strings when importing CSV data into a worksheet. | Provide an example of using CellsDataTableFactory to build a data table from a list of object arrays parsed from CSV.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Cells;

// Demonstrates a lightweight C# CSV parser that correctly handles commas, double quotes, and escaped quotes, builds an ICellsDataTable from the parsed rows (using the first row as a header), imports the table into a worksheet with numeric conversion and visible field names, and saves the result as an XLSX file.
class Program
{
    static void Main()
    {
        // Sample CSV content containing commas and quoted fields
        string csv = "Name,Age,Comment\n\"Doe, John\",30,\"He said, \"\"Hello!\"\"\"\n\"Smith, Jane\",25,\"New employee\"";

        // Parse the CSV into a list of object arrays (each array represents a row)
        List<object[]> rows = ParseCsv(csv);

        // Create a new workbook and get the first worksheet's cells
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Build an ICellsDataTable from the parsed rows.
        // The first row is treated as a header (hasHeader = true).
        ICellsDataTable dataTable = workbook.CellsDataTableFactory.GetInstance(
            rows.ToArray(),          // all rows including header
            true,                    // first row is header
            null);                   // column names are taken from the header row

        // Set import options: convert numeric strings to numbers and show field names.
        ImportTableOptions importOptions = new ImportTableOptions
        {
            ConvertNumericData = true,
            IsFieldNameShown = true
        };

        // Import the data table into the worksheet starting at cell A1 (row 0, column 0)
        cells.ImportData(dataTable, 0, 0, importOptions);

        // Save the workbook
        workbook.Save("CsvImported.xlsx", SaveFormat.Xlsx);
    }

    // Simple CSV parser that handles commas, double quotes and escaped quotes.
    static List<object[]> ParseCsv(string csvContent)
    {
        var result = new List<object[]>();
        using (StringReader reader = new StringReader(csvContent))
        {
            string line;
            while ((line = reader.ReadLine()) != null)
            {
                var fields = new List<object>();
                int i = 0;
                while (i < line.Length)
                {
                    if (line[i] == '\"')
                    {
                        // Quoted field
                        i++; // skip opening quote
                        var sb = new StringBuilder();
                        while (i < line.Length)
                        {
                            if (line[i] == '\"')
                            {
                                // Check for escaped double quote
                                if (i + 1 < line.Length && line[i + 1] == '\"')
                                {
                                    sb.Append('\"');
                                    i += 2;
                                }
                                else
                                {
                                    i++; // skip closing quote
                                    break;
                                }
                            }
                            else
                            {
                                sb.Append(line[i]);
                                i++;
                            }
                        }
                        fields.Add(sb.ToString());
                        // Skip delimiter if present
                        if (i < line.Length && line[i] == ',') i++;
                    }
                    else
                    {
                        // Unquoted field
                        int start = i;
                        while (i < line.Length && line[i] != ',') i++;
                        string token = line.Substring(start, i - start);
                        fields.Add(token);
                        if (i < line.Length && line[i] == ',') i++;
                    }
                }
                result.Add(fields.ToArray());
            }
        }
        return result;
    }
}
