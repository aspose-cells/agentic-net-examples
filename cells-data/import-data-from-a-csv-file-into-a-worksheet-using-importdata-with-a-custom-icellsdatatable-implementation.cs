// Title: Import CSV file into an Aspose.Cells worksheet using a custom ICellsDataTable in C#
// AI Prompts: Write C# code that reads a CSV file, parses numeric values, stores each row in an ArrayList, creates an ICellsDataTable via CellsDataTableFactory, and imports it into the first worksheet with ImportData. | Show how to build a custom ICellsDataTable from CSV data and save the resulting workbook as an XLSX file using Aspose.Cells. | Demonstrate configuring ImportTableOptions (default settings) when importing a custom data table into an Aspose.Cells worksheet.
// Common Searches: c# aspocells import csv using custom icellsdatatable | how to use CellsDataTableFactory to import csv data into an Excel workbook | import csv with numeric conversion to Aspose.Cells worksheet c# | Aspose.Cells ImportData from ArrayList example
// Tags: custom ICellsDataTable CSV import Aspose.Cells | CellsDataTableFactory usage for worksheet data loading | numeric parsing of CSV fields in C# Excel export | ImportTableOptions default configuration Aspose.Cells | ArrayList source for worksheet ImportData

using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCsvImportExample
{
    // The example reads a CSV file, splits each line into fields, converts numeric strings to int or double, stores rows in ArrayLists, creates a custom ICellsDataTable via Workbook.CellsDataTableFactory.GetInstance, imports the table into the first worksheet at cell A1 using Cells.ImportData, and saves the workbook as an XLSX file.
    class Program
    {
        static void Main()
        {
            // Path to the CSV file to be imported
            string csvPath = "data.csv";

            // Read all lines from the CSV file
            string[] lines = File.ReadAllLines(csvPath);

            // Prepare a collection where each item represents a row of the CSV.
            // Using ArrayList to match the expected type for CellsDataTableFactory.
            ArrayList dataLists = new ArrayList();

            foreach (string line in lines)
            {
                // Split the line by comma (you can change the delimiter if needed)
                string[] fields = line.Split(',');

                // Convert the string array to an object array and store it in an ArrayList
                ArrayList row = new ArrayList();
                foreach (string field in fields)
                {
                    // Try to parse numeric values; otherwise keep as string
                    if (int.TryParse(field, out int intVal))
                        row.Add(intVal);
                    else if (double.TryParse(field, out double doubleVal))
                        row.Add(doubleVal);
                    else
                        row.Add(field);
                }

                dataLists.Add(row);
            }

            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Build a custom ICellsDataTable from the CSV data (using the factory)
            ICellsDataTable dataTable = workbook.CellsDataTableFactory.GetInstance(dataLists, true);

            // Import the custom data table into the first worksheet starting at cell A1
            // ImportTableOptions can be customized; using defaults here
            workbook.Worksheets[0].Cells.ImportData(dataTable, 0, 0, new ImportTableOptions());

            // Save the workbook (lifecycle: save)
            workbook.Save("ImportedFromCsv.xlsx", SaveFormat.Xlsx);

            Console.WriteLine("CSV data has been imported and saved to 'ImportedFromCsv.xlsx'.");
        }
    }
}
