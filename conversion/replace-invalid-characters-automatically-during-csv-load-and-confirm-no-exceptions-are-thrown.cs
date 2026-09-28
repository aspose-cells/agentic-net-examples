// Title: Sanitize CSV by removing illegal control characters and load it into an Aspose.Cells workbook without errors (C#)
// AI Prompts: Generate C# code that reads a CSV file, strips any characters outside the allowed Unicode range using Regex, and creates an Aspose.Cells Workbook from a MemoryStream. | Provide a reusable C# method named CleanCsvForAspose that takes raw CSV text, removes prohibited control characters, and returns a Workbook instance. | Show how to extend the CSV sanitization to replace embedded line‑breaks inside quoted fields before loading the data with Aspose.Cells.
// Common Searches: C# how to filter out illegal control characters from a CSV before using Aspose.Cells LoadOptions | Aspose.Cells CSV import throws exception due to hidden characters, how to prevent it | Regex pattern to keep only allowed Unicode characters when loading CSV into Aspose.Cells | Load CSV into Aspose.Cells workbook from memory stream after cleaning content
// Tags: regex sanitization of CSV for Aspose.Cells | load cleaned CSV via MemoryStream in C# | remove illegal control characters before Aspose.Cells import | Aspose.Cells CSV load options with sanitized input | C# CSV preprocessing for Excel workbook conversion

using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example reads a CSV file, uses a regular expression to delete characters that are not permitted in CSV (all control characters except TAB, LF, CR), loads the cleaned text into an Aspose.Cells Workbook via a MemoryStream, and saves it as XLSX, ensuring no load‑time exceptions.
class CsvLoader
{
    static void Main()
    {
        // Paths for input CSV and output workbook
        string inputCsvPath = "input.csv";
        string outputWorkbookPath = "output.xlsx";

        // Read the entire CSV file as text
        string csvText = File.ReadAllText(inputCsvPath, Encoding.UTF8);

        // Replace characters that are not allowed in CSV (control chars except TAB, LF, CR)
        // This ensures no exception is thrown during load.
        string cleanedCsv = Regex.Replace(csvText, @"[^\u0009\u000A\u000D\u0020-\uFFFF]", string.Empty);

        // Load the cleaned CSV content into a workbook using a memory stream
        using (MemoryStream csvStream = new MemoryStream(Encoding.UTF8.GetBytes(cleanedCsv)))
        {
            // LoadOptions specify that the source format is CSV
            LoadOptions loadOptions = new LoadOptions(LoadFormat.CSV);

            // Create workbook from the cleaned CSV stream
            Workbook workbook = new Workbook(csvStream, loadOptions);

            // Save the workbook to the desired format (e.g., XLSX)
            workbook.Save(outputWorkbookPath, SaveFormat.Xlsx);
        }

        Console.WriteLine("CSV loaded successfully and saved as workbook.");
    }
}
