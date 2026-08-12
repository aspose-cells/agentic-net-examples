// Title: C# – Convert JSON to Pipe‑Delimited CSV with Aspose.Cells JsonUtility
// Description: A complete C# example that reads a JSON file, uses Aspose.Cells JsonUtility with the ArrayAsTable layout to import the data into a worksheet, then iterates the used range to generate a CSV file using a custom "|" delimiter. The code handles delimiter/quote escaping, writes the CSV to disk, and optionally saves an intermediate XLSX for verification.
// Keywords: Aspose.Cells JsonUtility | C# JSON to CSV | pipe delimited CSV | ArrayAsTable | custom delimiter export | JSON import Excel worksheet | Aspose.Cells CSV export | C# data conversion
// Common Searches: Aspose.Cells JsonUtility export JSON as CSV with custom delimiter | C# convert JSON file to pipe‑separated values using Aspose.Cells | How to use ArrayAsTable option for JSON to CSV conversion in C# | Save intermediate Excel workbook while converting JSON to CSV
// Developer Intent: Create a pipe‑delimited CSV from a JSON source by importing the JSON into an Aspose.Cells worksheet and exporting the cell values.
// Use Cases: Generate legacy‑system reports that require a non‑standard "|" separator. | Validate JSON‑derived data visually in Excel before distributing CSV files. | Integrate JSON API responses into data‑pipeline workflows that consume custom‑delimited CSV.
// AI Prompts: Write C# code that reads a JSON file, imports it into an Aspose.Cells worksheet with ArrayAsTable, and exports a semicolon‑delimited CSV, handling proper escaping. | Refactor the example to stream a large JSON file to CSV using Aspose.Cells without loading the entire file into memory. | Show how to add column headers from JSON property names when exporting to a custom‑delimited CSV with Aspose.Cells.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonToCsv
{
    // A complete C# example that reads a JSON file, uses Aspose.Cells JsonUtility with the ArrayAsTable layout to import the data into a worksheet, then iterates the used range to generate a CSV file using a custom "|" delimiter. The code handles delimiter/quote escaping, writes the CSV to disk, and optionally saves an intermediate XLSX for verification.
    class Program
    {
        static void Main()
        {
            // Path to the source JSON file
            string jsonPath = "data.json";

            // Custom delimiter for CSV output
            string delimiter = "|";

            // Read the entire JSON content
            string jsonContent = File.ReadAllText(jsonPath);

            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet and its cells
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Set JSON layout options to treat arrays as tables
            JsonLayoutOptions layoutOptions = new JsonLayoutOptions
            {
                ArrayAsTable = true
            };

            // Import JSON data into the worksheet starting at cell A1
            JsonUtility.ImportData(jsonContent, cells, 0, 0, layoutOptions);

            // Determine the used range of the worksheet
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            // Build CSV content using the custom delimiter
            StringBuilder csvBuilder = new StringBuilder();

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    // Get cell value as string; handle nulls
                    string cellText = cells[row, col]?.StringValue ?? string.Empty;

                    // Escape delimiter and quotes if necessary
                    if (cellText.Contains(delimiter) || cellText.Contains("\""))
                    {
                        cellText = $"\"{cellText.Replace("\"", "\"\"")}\"";
                    }

                    csvBuilder.Append(cellText);

                    // Append delimiter except after the last column
                    if (col < maxCol)
                        csvBuilder.Append(delimiter);
                }

                // New line after each row
                csvBuilder.AppendLine();
            }

            // Write the CSV content to a file
            string csvPath = "output.csv";
            File.WriteAllText(csvPath, csvBuilder.ToString(), Encoding.UTF8);

            // Optionally, save the workbook as an Excel file for verification
            workbook.Save("intermediate.xlsx", SaveFormat.Xlsx);

            Console.WriteLine($"JSON data has been converted to CSV with delimiter '{delimiter}' and saved to '{csvPath}'.");
        }
    }
}
