// Title: Handle and Log CSV Parsing Errors with a Custom ICustomParser in Aspose.Cells for .NET
// Description: Demonstrates a StringParser that implements ICustomParser and throws an exception when a cell contains "ERROR". Shows how to configure TxtLoadOptions, load CSV data inside a try‑catch block, log parsing failures, access cells, and save the workbook as an Excel file.
// Keywords: Aspose.Cells | CSV import | custom parser | ICustomParser | error handling | parsing exception | TxtLoadOptions | .NET | C# | workbook save | log CSV errors
// Common Searches: Aspose.Cells catch CSV parsing exception | log errors when loading CSV with custom parser Aspose.Cells | ICustomParser example with exception handling | load CSV to Excel workbook with error handling .NET | skip bad rows during CSV import Aspose.Cells
// Developer Intent: The developer needs to capture and record any parsing exceptions that occur while loading a CSV file using custom parsers in Aspose.Cells.
// Use Cases: Identify and log malformed rows during CSV import to maintain data integrity | Continue processing remaining rows after a parsing failure | Provide clear error messages for end‑users when custom validation fails | Integrate CSV import errors with monitoring or alerting systems
// AI Prompts: Generate C# code that wraps Aspose.Cells CSV loading in a try‑catch block and writes detailed exception information to a log file using Serilog. | Show how to modify StringParser so it records the offending value and allows the load operation to skip the problematic row. | Provide an example of an ICustomParser that validates numeric cells and gracefully handles conversion errors without stopping the import.

using System;
using System.IO;
using Aspose.Cells;

// Custom parser that throws an exception for values containing "ERROR"
// Demonstrates a StringParser that implements ICustomParser and throws an exception when a cell contains "ERROR". Shows how to configure TxtLoadOptions, load CSV data inside a try‑catch block, log parsing failures, access cells, and save the workbook as an Excel file.
class StringParser : ICustomParser
{
    public object ParseObject(string value)
    {
        if (value.Contains("ERROR"))
            throw new Exception($"Parsing error for value: {value}");
        return value; // Return the original string if no error
    }

    public string GetFormat()
    {
        return "String";
    }
}

class Program
{
    static void Main()
    {
        // Sample CSV data; the second row contains a value that will cause a parsing exception
        string csvData = "Name,Value\nJohn,123\nBadRow,ERROR_VALUE";

        // Convert CSV string to a memory stream
        using (MemoryStream csvStream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(csvData)))
        {
            // Configure TxtLoadOptions with a custom parser for the first column
            TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv);
            loadOptions.PreferredParsers = new ICustomParser[] { new StringParser(), null };

            Workbook workbook = null;

            try
            {
                // Load the CSV into a workbook using the configured options
                workbook = new Workbook(csvStream, loadOptions);
                Console.WriteLine("CSV loaded successfully.");
            }
            catch (Exception ex)
            {
                // Log any parsing exceptions that occur during loading
                Console.WriteLine($"Error loading CSV: {ex.Message}");
            }

            if (workbook != null)
            {
                // Access the loaded cells (demonstration)
                Cells cells = workbook.Worksheets[0].Cells;
                Console.WriteLine($"A1: {cells["A1"].StringValue}");
                Console.WriteLine($"B2: {cells["B2"].StringValue}");

                // Save the workbook to an Excel file
                workbook.Save("Output.xlsx", SaveFormat.Xlsx);
                Console.WriteLine("Workbook saved as Output.xlsx.");
            }
        }
    }
}
