// Title: Import CSV into an Aspose.Cells worksheet using a custom ICellsDataTable (C#)
// Description: Read a CSV file line‑by‑line, convert each field to int, double or string, build an ArrayList of object arrays, create an ICellsDataTable with CellsDataTableFactory, import it to the first worksheet at A1 via ImportData, and save as XLSX.
// Keywords: Aspose.Cells ImportData | ICellsDataTable | CellsDataTableFactory | CSV to Excel C# | type conversion CSV Aspose | ImportTableOptions | Workbook.Save XLSX | .NET Excel library
// Common Searches: how to import csv into Aspose.Cells worksheet | custom ICellsDataTable example C# | ImportData with ImportTableOptions Aspose | convert csv values to int double Aspose.Cells | read csv into workbook using CellsDataTableFactory
// Developer Intent: Load CSV content into an Excel workbook while preserving native .NET data types via a custom ICellsDataTable.
// Use Cases: Parse a CSV file, infer numeric types, and populate an Excel sheet for reporting. | Create a reusable ICellsDataTable from any IEnumerable of object arrays for bulk import. | Apply ImportTableOptions (e.g., header handling, column width auto‑fit) during the import process. | Generate an XLSX file that can be further processed or shared after CSV ingestion.
// AI Prompts: Show how to skip the first row (header) when importing CSV with ICellsDataTable. | Provide C# code that uses a DataTable instead of an ArrayList with CellsDataTableFactory. | Explain which ImportTableOptions settings preserve formulas or formatting during CSV import.

using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

// Read a CSV file line‑by‑line, convert each field to int, double or string, build an ArrayList of object arrays, create an ICellsDataTable with CellsDataTableFactory, import it to the first worksheet at A1 via ImportData, and save as XLSX.
class ImportCsvWithCustomDataTable
{
    static void Main()
    {
        // Path to the CSV file to be imported
        string csvPath = "data.csv";

        // Read the CSV file line by line and build a collection of object arrays
        ArrayList dataLists = new ArrayList();
        using (StreamReader reader = new StreamReader(csvPath))
        {
            string line;
            while ((line = reader.ReadLine()) != null)
            {
                // Split each line by comma (adjust delimiter if needed)
                string[] parts = line.Split(',');

                // Convert each field to the most appropriate type (int, double, or string)
                object[] row = new object[parts.Length];
                for (int i = 0; i < parts.Length; i++)
                {
                    string value = parts[i].Trim();

                    if (int.TryParse(value, out int intVal))
                        row[i] = intVal;
                    else if (double.TryParse(value, out double doubleVal))
                        row[i] = doubleVal;
                    else
                        row[i] = value;
                }

                dataLists.Add(row);
            }
        }

        // Create a new workbook
        Workbook workbook = new Workbook();

        // Build an ICellsDataTable from the collection using the factory
        ICellsDataTable cellsTable = workbook.CellsDataTableFactory.GetInstance(dataLists, true);

        // Import the data table into the first worksheet starting at cell A1 (row 0, column 0)
        workbook.Worksheets[0].Cells.ImportData(cellsTable, 0, 0, new ImportTableOptions());

        // Save the workbook to an XLSX file
        workbook.Save("ImportedFromCsv.xlsx");
    }
}
