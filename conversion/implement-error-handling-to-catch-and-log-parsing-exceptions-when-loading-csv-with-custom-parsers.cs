// Title: Catch and log parsing exceptions when loading a CSV with a column‑specific custom parser using Aspose.Cells for .NET
// AI Prompts: Create C# code that loads CSV data into an Aspose.Cells Workbook with TxtLoadOptions, assigns a custom ICustomParser to a chosen column, and surrounds the load operation with try‑catch to log any parsing errors. | Update an existing Aspose.Cells CSV import routine to record the row and column of a value that triggers a FormatException from a custom parser, then decide whether to continue or abort based on the logged information.
// Common Searches: asp.net how to catch formatexception from a custom csv parser in Aspose.Cells | Aspose.Cells load csv using TxtLoadOptions PreferredParsers and log parsing failures | error handling for column specific parser when converting csv to xlsx with Aspose.Cells | log faulty csv values during workbook import in .NET using Aspose.Cells
// Tags: custom column parser Aspose.Cells | csv load exception logging .NET | TxtLoadOptions PreferredParsers example | Aspose.Cells FormatException capture | csv to xlsx conversion with fault tolerance

using System;
using System.IO;
using Aspose.Cells;

// The example defines a FaultyParser that implements ICustomParser and throws a FormatException for values containing "ERR". It assigns this parser to the second column via TxtLoadOptions.PreferredParsers, loads CSV data from a memory stream into a Workbook, saves the workbook to XLSX on success, and catches any exceptions during loading to log the error message.
class CsvLoaderWithErrorHandling
{
    // Custom parser that throws an exception for values containing "ERR"
    private class FaultyParser : ICustomParser
    {
        public bool Parse(string value, out object result)
        {
            if (value.Contains("ERR"))
                throw new FormatException($"Unable to parse value '{value}'.");
            if (double.TryParse(value, out double d))
            {
                result = d;
                return true;
            }
            result = value;
            return true;
        }

        public object ParseObject(string value)
        {
            // Not used in this scenario
            return value;
        }

        public string GetFormat()
        {
            return "Custom";
        }
    }

    static void Main()
    {
        // Sample CSV data with an intentional parsing error in the second column of the second row
        string csvData = "Name,Score\nAlice,85\nBob,ERR\nCharlie,92";
        byte[] bytes = System.Text.Encoding.UTF8.GetBytes(csvData);

        using (MemoryStream stream = new MemoryStream(bytes))
        {
            // Configure load options: use the custom parser for the second column
            TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv);
            loadOptions.PreferredParsers = new ICustomParser[] { null, new FaultyParser() };

            try
            {
                // Load the workbook with the custom parsers; any parsing exception will be caught
                Workbook workbook = new Workbook(stream, loadOptions);

                // Save the workbook if loading succeeded
                workbook.Save("Result.xlsx", SaveFormat.Xlsx);
                Console.WriteLine("CSV loaded and saved successfully.");
            }
            catch (Exception ex)
            {
                // Log parsing exception details
                Console.WriteLine($"Error loading CSV: {ex.Message}");
            }
        }
    }
}
