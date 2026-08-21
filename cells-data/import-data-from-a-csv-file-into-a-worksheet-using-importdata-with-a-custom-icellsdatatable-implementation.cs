// Title: C# – Import CSV into an Aspose.Cells worksheet with a custom ICellsDataTable
// Description: Read a CSV file, split each line into an object array, build an ICellsDataTable via CellsDataTableFactory, and import the data into the first worksheet using Cells.ImportData. The workbook is then saved as an XLSX file.
// Keywords: Aspose.Cells | C# | CSV import | ICellsDataTable | ImportData | CellsDataTableFactory | ArrayList | Workbook | ImportTableOptions | Excel generation
// Common Searches: Aspose.Cells import CSV C# example | How to use ICellsDataTable with ImportData | Create custom ICellsDataTable from collection | Import CSV data into Aspose.Cells worksheet | C# read CSV and load into Excel using Aspose
// Developer Intent: Load CSV content into an Excel worksheet by converting rows to a custom ICellsDataTable and calling ImportData.
// Use Cases: Convert a comma‑delimited CSV file into an Excel workbook with a single ImportData call. | Reuse the same ICellsDataTable to populate multiple worksheets in one workbook. | Import CSV data while optionally skipping the header row or customizing column widths via ImportTableOptions.
// AI Prompts: Generate C# code that reads a semicolon‑delimited CSV, builds an ICellsDataTable, and imports it using custom ImportTableOptions for column widths. | Show how to exclude the first row (header) when creating the ICellsDataTable and then import the remaining rows. | Provide a snippet that imports the same CSV data into three worksheets, each starting at a different cell address.

using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCsvImport
{
    // Read a CSV file, split each line into an object array, build an ICellsDataTable via CellsDataTableFactory, and import the data into the first worksheet using Cells.ImportData. The workbook is then saved as an XLSX file.
    class Program
    {
        static void Main()
        {
            // Path to the CSV file to be imported
            string csvPath = "data.csv";

            // Read all lines from the CSV file
            string[] lines = File.ReadAllLines(csvPath);

            // Prepare a collection that will hold each row as an object array
            ArrayList dataLists = new ArrayList();

            foreach (string line in lines)
            {
                // Split the line by comma (you can change the delimiter if needed)
                string[] parts = line.Split(',');

                // Convert the string parts to an object array and add to the collection
                dataLists.Add(parts);
            }

            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Build an ICellsDataTable from the collection (custom data table)
            ICellsDataTable dataTable = workbook.CellsDataTableFactory.GetInstance(dataLists, true);

            // Import the data table into the first worksheet starting at cell A1
            // ImportTableOptions can be customized; using defaults here
            workbook.Worksheets[0].Cells.ImportData(dataTable, 0, 0, new ImportTableOptions());

            // Save the workbook (lifecycle rule: save)
            workbook.Save("ImportedFromCsv.xlsx");
        }
    }
}
