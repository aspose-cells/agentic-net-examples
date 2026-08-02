// Title: C# – Load CSV into Aspose.Cells Workbook while preserving numeric values as text (ConvertNumericToText)
// Description: Demonstrates how to use TxtLoadOptions with ConvertNumericData = false (ConvertNumericToText) to load CSV data from a stream into an Aspose.Cells Workbook, keeping numeric cells as text and showing their types before optionally saving to XLSX.
// Keywords: Aspose.Cells load CSV as text | ConvertNumericToText option | TxtLoadOptions numeric conversion false | read numeric cells as string C# | preserve numeric formatting Aspose.Cells
// Common Searches: Aspose.Cells prevent numeric conversion CSV | TxtLoadOptions ConvertNumericData example | load CSV keep numbers as text Aspose.Cells | C# read numbers as strings in workbook | Aspose.Cells CSV import text values
// Developer Intent: Load a CSV file into an Aspose.Cells Workbook while ensuring that numeric columns are treated as text strings.
// Use Cases: Import CSV streams where leading zeros or exact formatting must be retained. | Validate cell data types after import to confirm numeric values were not auto‑converted. | Save the workbook for downstream Excel reporting while preserving original text representation of numbers.
// AI Prompts: Show C# code that loads a CSV into Aspose.Cells and disables numeric conversion using ConvertNumericToText. | Provide an example that reads a CSV with TxtLoadOptions, prints each cell's type and string value, and saves the workbook. | Explain how to configure TxtLoadOptions to keep numeric columns as text when importing data into a Workbook.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// Demonstrates how to use TxtLoadOptions with ConvertNumericData = false (ConvertNumericToText) to load CSV data from a stream into an Aspose.Cells Workbook, keeping numeric cells as text and showing their types before optionally saving to XLSX.
class ReadNumericAsTextDemo
{
    static void Main()
    {
        // Sample CSV content containing numeric values
        string csvData = "ID,Price,Quantity\n1,19.99,5\n2,24.50,10";

        // Configure load options for CSV and disable numeric conversion
        TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv)
        {
            ConvertNumericData = false // keep numeric values as text
        };

        // Load the CSV data into a workbook using the configured options
        using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(csvData)))
        {
            Workbook workbook = new Workbook(stream, loadOptions);
            Worksheet sheet = workbook.Worksheets[0];

            // Demonstrate that numeric cells are treated as text
            Console.WriteLine($"Cell B2 Type: {sheet.Cells["B2"].Type}, Value: {sheet.Cells["B2"].StringValue}");
            Console.WriteLine($"Cell C2 Type: {sheet.Cells["C2"].Type}, Value: {sheet.Cells["C2"].StringValue}");

            // Save the workbook (optional)
            workbook.Save("Output.xlsx");
        }
    }
}
