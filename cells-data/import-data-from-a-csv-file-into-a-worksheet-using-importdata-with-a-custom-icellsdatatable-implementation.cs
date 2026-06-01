using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCsvImport
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the CSV file to be imported
                string csvPath = "data.csv";

                // Verify that the CSV file exists
                if (!File.Exists(csvPath))
                {
                    Console.WriteLine($"Error: CSV file not found at path '{csvPath}'.");
                    return;
                }

                // Read all lines from the CSV file
                string[] lines = File.ReadAllLines(csvPath);

                // Prepare a collection where each item represents a row of the CSV file
                // Aspose.Cells expects an ArrayList of rows, where each row is an ArrayList of column values
                ArrayList dataRows = new ArrayList();

                foreach (string line in lines)
                {
                    // Split the line by comma (you can change the delimiter if needed)
                    string[] fields = line.Split(',');

                    // Create a row as an ArrayList and add the fields
                    ArrayList row = new ArrayList();
                    row.AddRange(fields);

                    // Add the row to the collection
                    dataRows.Add(row);
                }

                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Obtain the factory for building ICellsDataTable from the workbook
                CellsDataTableFactory factory = workbook.CellsDataTableFactory;

                // Build an ICellsDataTable from the prepared collection
                // The second parameter 'true' indicates that the first row contains column names
                ICellsDataTable dataTable = factory.GetInstance(dataRows, true);

                // Import the custom data table into the first worksheet starting at cell A1 (row 0, column 0)
                ImportTableOptions importOptions = new ImportTableOptions();
                workbook.Worksheets[0].Cells.ImportData(dataTable, 0, 0, importOptions);

                // Save the workbook to an XLSX file (lifecycle rule: save)
                string outputPath = "ImportedFromCsv.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"CSV data has been imported successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}