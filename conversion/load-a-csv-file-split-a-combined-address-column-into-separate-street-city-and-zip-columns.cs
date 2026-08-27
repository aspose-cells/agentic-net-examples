// Title: Import a CSV with a combined address field and split it into Street, City, and Zip columns using Aspose.Cells TextToColumns in C#
// AI Prompts: Load a CSV file into an Aspose.Cells Workbook, configure TxtLoadOptions with a comma separator, apply the TextToColumns method to separate the address column into Street, City, and Zip, then save the workbook as an XLSX file. | Demonstrate creating a sample CSV, importing it via Workbook.Cells.ImportCSV, using TextToColumns on column A to split the address, and exporting the resulting worksheet to Excel with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# split address column from CSV into separate columns | How to use TextToColumns to parse combined address in a CSV with Aspose.Cells | Import CSV and separate street city zip using Aspose.Cells TextToColumns method | C# Aspose.Cells example for splitting address field into multiple columns | Convert CSV address list to Excel with street city zip columns using Aspose.Cells
// Tags: Aspose.Cells CSV import and address column split | Configure TxtLoadOptions separator for TextToColumns | C# TextToColumns on first worksheet column | Save split address workbook as XLSX | Workbook.Cells.ImportCSV example in .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AddressSplitExample
{
    // The program creates a Workbook, imports a CSV containing a combined address field, sets TxtLoadOptions with a comma separator, uses TextToColumns to divide the address into Street, City, and Zip columns, and saves the result as an XLSX file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the source CSV file that contains a combined address column (e.g., "123 Main St,City,12345")
                string csvPath = "addresses.csv";

                // Ensure the CSV file exists; create a simple sample if it does not.
                if (!File.Exists(csvPath))
                {
                    string sampleData = "Address\n\"123 Main St,CityA,12345\"\n\"456 Oak Rd,CityB,67890\"";
                    File.WriteAllText(csvPath, sampleData);
                }

                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Access the first worksheet and its cells collection
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Import the CSV data starting at cell A1 (row 0, column 0)
                // Using comma as the delimiter and converting numeric data where possible
                cells.ImportCSV(csvPath, ",", true, 0, 0); // lifecycle rule: load (ImportCSV)

                // Determine how many rows were imported (including header if present)
                int totalRows = sheet.Cells.MaxDataRow + 1;

                // Configure split options: use comma as the separator for the address column
                TxtLoadOptions splitOptions = new TxtLoadOptions
                {
                    Separator = ',' // or splitOptions.SeparatorString = ","
                };

                // Split the combined address column (column A, index 0) into separate columns:
                // Street -> column A, City -> column B, Zip -> column C
                cells.TextToColumns(0, 0, totalRows, splitOptions); // lifecycle rule: split (TextToColumns)

                // Save the resulting workbook to an XLSX file (lifecycle rule: save)
                workbook.Save("addresses_split.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
