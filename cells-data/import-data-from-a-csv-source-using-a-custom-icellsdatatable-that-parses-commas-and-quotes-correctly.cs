using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCsvImportExample
{
    class Program
    {
        // Simple CSV parser that handles commas and double quotes according to RFC 4180
        static List<object[]> ParseCsv(string csvContent)
        {
            var rows = new List<object[]>();
            using (StringReader sr = new StringReader(csvContent))
            {
                string line;
                while ((line = sr.ReadLine()) != null)
                {
                    var fields = new List<string>();
                    bool inQuotes = false;
                    var field = string.Empty;

                    for (int i = 0; i < line.Length; i++)
                    {
                        char c = line[i];

                        if (c == '\"')
                        {
                            // If double quote inside quoted field, check for escaped quote
                            if (inQuotes && i + 1 < line.Length && line[i + 1] == '\"')
                            {
                                field += '\"';
                                i++; // skip escaped quote
                            }
                            else
                            {
                                inQuotes = !inQuotes; // toggle quoting state
                            }
                        }
                        else if (c == ',' && !inQuotes)
                        {
                            fields.Add(field);
                            field = string.Empty;
                        }
                        else
                        {
                            field += c;
                        }
                    }
                    // Add last field
                    fields.Add(field);

                    // Convert to object array (keep as string; ImportData will convert numeric if needed)
                    rows.Add(fields.ToArray());
                }
            }
            return rows;
        }

        static void Main(string[] args)
        {
            // Sample CSV content with commas and quoted fields
            string csvData = 
                "Name,Age,Comment\n" +
                "\"Doe, John\",30,\"Works at \"\"Acme, Inc.\"\"\"\n" +
                "Smith,25,\"New employee\"\n" +
                "\"Brown, \"\"Bob\"\"\",40,\"Senior, Manager\"";

            // Parse CSV into a collection of object arrays
            List<object[]> parsedRows = ParseCsv(csvData);

            // Convert List<object[]> to a non‑generic ICollection for the factory
            ICollection dataCollection = new ArrayList(parsedRows);

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Obtain the CellsDataTableFactory from the workbook
            CellsDataTableFactory factory = workbook.CellsDataTableFactory;

            // Create ICellsDataTable from the collection; true indicates vertical orientation (rows)
            ICellsDataTable dataTable = factory.GetInstance(dataCollection, true);

            // Import the data table into the first worksheet starting at cell A1
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Cells.ImportData(dataTable, 0, 0, new ImportTableOptions());

            // Save the workbook to an XLSX file
            workbook.Save("CsvImportedWithCustomDataTable.xlsx", SaveFormat.Xlsx);
        }
    }
}