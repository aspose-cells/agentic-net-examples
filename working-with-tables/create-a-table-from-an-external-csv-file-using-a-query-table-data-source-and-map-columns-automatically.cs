// Title: Create an Excel Table from a CSV File with Automatic Column Mapping using Aspose.Cells (C#)
// Description: Demonstrates how to verify a CSV file, import its data into a new workbook, generate a ListObject (Excel table) with headers, auto‑fit the columns, and save the result as an XLSX file using Aspose.Cells in C#.
// Keywords: Aspose.Cells CSV import | C# ImportCSV to ListObject | Excel table from CSV | automatic column mapping Aspose.Cells | auto fit columns C# | save workbook as XLSX | query table Aspose.Cells | ListObject example C# | Aspose.Cells data import | CSV to Excel table conversion
// Common Searches: Aspose.Cells import CSV and create table C# | How to add ListObject from CSV using Aspose.Cells | C# auto‑fit columns after CSV import Aspose | Create Excel table from external CSV file Aspose.Cells | Check CSV file existence before importing Aspose
// Developer Intent: Read a CSV file, convert it into an Excel table with mapped columns, and export the workbook as XLSX.
// Use Cases: Generate a formatted report by turning raw CSV data into an Excel table with headers. | Encapsulate CSV‑to‑table conversion in a reusable C# method for batch processing. | Prepare data for downstream analysis by importing CSV, creating a ListObject, and applying auto‑fit.
// AI Prompts: Write a C# function that takes a CSV path, imports the data with Aspose.Cells, creates a ListObject with headers, auto‑fits columns, and returns the Workbook. | Provide error‑handling code for missing CSV files and invalid delimiters when building a query table with Aspose.Cells. | Show how to add a query table to an existing worksheet that references an external CSV and automatically maps its columns using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsQueryTableFromCsv
{
    // Demonstrates how to verify a CSV file, import its data into a new workbook, generate a ListObject (Excel table) with headers, auto‑fit the columns, and save the result as an XLSX file using Aspose.Cells in C#.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the external CSV file
                string csvPath = "data.csv";

                // Verify that the CSV file exists to avoid FileNotFoundException
                if (!File.Exists(csvPath))
                {
                    Console.WriteLine($"CSV file not found: {csvPath}");
                    return;
                }

                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // 2. Import the CSV data into the worksheet starting at cell A1 (row 0, column 0)
                //    Use comma as the delimiter and enable numeric conversion.
                cells.ImportCSV(csvPath, ",", true, 0, 0);

                // 3. Determine the range that now contains data
                //    MaxDisplayRange gives the smallest rectangle that encloses all non‑empty cells.
                AsposeRange dataRange = cells.MaxDisplayRange;

                // 4. Add a ListObject (Excel table) over the imported data.
                //    The parameters are: first row, first column, total rows, total columns, hasHeaders.
                int listObjectIndex = sheet.ListObjects.Add(
                    dataRange.FirstRow,          // first row
                    dataRange.FirstColumn,       // first column
                    dataRange.RowCount,          // total rows
                    dataRange.ColumnCount,       // total columns
                    true);                       // show header

                ListObject table = sheet.ListObjects[listObjectIndex];

                // 5. (Optional) Adjust column widths automatically based on the imported data.
                //    AutoFitColumns works on the worksheet; specify the column range of the table.
                sheet.AutoFitColumns(dataRange.FirstColumn, dataRange.ColumnCount);

                // 6. Save the workbook to an XLSX file.
                workbook.Save("ResultFromCsv.xlsx", SaveFormat.Xlsx);
                Console.WriteLine("Workbook saved as ResultFromCsv.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
