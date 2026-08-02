// Title: C# – Catch and Log CSV Parsing Errors with a Custom ICustomParser in Aspose.Cells
// Description: Shows how to configure TxtLoadOptions with a custom ICustomParser, load CSV data, catch CellsException (InvalidData) and other exceptions, log the problem, and save the workbook only when the import succeeds.
// Keywords: Aspose.Cells CSV custom parser | ICustomParser error handling | TxtLoadOptions parsing exception | CellsException InvalidData C# | log CSV import errors Aspose | load CSV with try‑catch Aspose.Cells | C# CSV to XLSX conversion error handling
// Common Searches: Aspose.Cells catch parsing errors CSV | C# custom parser for CSV with Aspose | how to log CellsException InvalidData | try‑catch workbook load Aspose.Cells | handle format exception during CSV import
// Developer Intent: Add robust try‑catch logic around Workbook loading to capture and record parsing failures when a custom ICustomParser encounters invalid data.
// Use Cases: Identify and log rows that contain prohibited tokens while importing CSV files. | Prevent workbook creation when the CSV contains format errors, ensuring data integrity. | Continue processing after logging a parsing error by skipping the problematic row.
// AI Prompts: Generate C# code that wraps Aspose.Cells Workbook loading with TxtLoadOptions in a try‑catch block, logs CellsException InvalidData to a file, and proceeds with the next row. | Create an ICustomParser that validates numeric ranges, throws a FormatException on out‑of‑range values, and demonstrate handling those exceptions during CSV import.

using System;
using System.IO;
using Aspose.Cells;

// Shows how to configure TxtLoadOptions with a custom ICustomParser, load CSV data, catch CellsException (InvalidData) and other exceptions, log the problem, and save the workbook only when the import succeeds.
class CsvLoaderWithErrorHandling
{
    // Custom parser implementing ICustomParser
    private class SafeStringParser : ICustomParser
    {
        public bool Parse(string value, out object result)
        {
            // Not used in this scenario; always succeed
            result = value;
            return true;
        }

        public object ParseObject(string value)
        {
            // Simulate a parsing error for a specific token
            if (value == "ERROR")
                throw new FormatException("Invalid value encountered during parsing.");
            return value;
        }

        public string GetFormat()
        {
            return "String";
        }
    }

    static void Main()
    {
        // Sample CSV data containing a deliberately bad value to trigger an exception
        string csvData = "\"Good\",123\n\"ERROR\",456";

        // Load CSV data into a memory stream
        using (MemoryStream csvStream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(csvData)))
        {
            // Configure TxtLoadOptions with a custom parser for the first column
            TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv);
            loadOptions.PreferredParsers = new ICustomParser[] { new SafeStringParser(), null };
            loadOptions.ConvertNumericData = true; // Allow numeric conversion for other columns

            Workbook workbook = null;
            try
            {
                // Attempt to load the CSV using the custom options
                workbook = new Workbook(csvStream, loadOptions);
                Console.WriteLine("CSV loaded successfully.");
            }
            catch (CellsException ex) when (ex.Code == ExceptionType.InvalidData)
            {
                // Handle parsing-specific errors
                Console.WriteLine($"Parsing error: {ex.Message}");
            }
            catch (Exception ex)
            {
                // Handle any other loading errors
                Console.WriteLine($"Error loading CSV: {ex.Message}");
            }

            // If loading succeeded, save the workbook
            if (workbook != null)
            {
                workbook.Save("LoadedCsv.xlsx", SaveFormat.Xlsx);
                Console.WriteLine("Workbook saved as LoadedCsv.xlsx");
            }
        }
    }
}
