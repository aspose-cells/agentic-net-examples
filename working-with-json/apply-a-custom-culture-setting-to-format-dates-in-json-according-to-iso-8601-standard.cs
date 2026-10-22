// Title: How to apply a custom ISO‑8601 CultureInfo to format dates when saving an Excel workbook as JSON with Aspose.Cells for .NET
// AI Prompts: Generate C# code that clones InvariantCulture, sets ISO‑8601 date patterns, assigns it to Workbook.Settings.CultureInfo, and saves the workbook to JSON. | Show the steps to configure a custom CultureInfo for ISO‑8601 date formatting in Aspose.Cells before calling Workbook.Save with SaveFormat.Json. | Provide an example that creates a workbook, applies ISO‑8601 short and long date/time patterns via CultureInfo, and exports the data to a JSON file.
// Common Searches: Aspose.Cells set ISO 8601 date format for JSON export in C# | C# save Excel workbook to JSON with dates in yyyy-MM-ddTHH:mm:ss | How to change workbook culture to ISO 8601 before JSON serialization using Aspose.Cells | Custom CultureInfo for date formatting when using Aspose.Cells SaveFormat.Json | Export Excel to JSON with ISO date strings using Aspose.Cells .NET
// Tags: Workbook.Settings.CultureInfo ISO 8601 | Aspose.Cells JSON date formatting | custom CultureInfo for JSON export | C# Excel to JSON with ISO dates | SaveFormat.Json date pattern customization

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// The example loads or creates an Excel workbook, clones the invariant culture, sets its date and time patterns to ISO‑8601 formats, assigns this CultureInfo to the workbook's Settings, and then saves the workbook as a JSON file so all dates appear in the ISO‑8601 style.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.json";

            // Load existing workbook or create a new one if the file is missing
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                // Example data to demonstrate date formatting
                workbook.Worksheets[0].Cells["A1"].PutValue(DateTime.Now);
            }

            // Configure ISO 8601 date/time patterns via a custom culture
            CultureInfo isoCulture = (CultureInfo)CultureInfo.InvariantCulture.Clone();
            isoCulture.DateTimeFormat.FullDateTimePattern = "yyyy-MM-ddTHH:mm:ss";
            isoCulture.DateTimeFormat.ShortDatePattern = "yyyy-MM-dd";
            isoCulture.DateTimeFormat.LongTimePattern = "HH:mm:ss";

            // Apply the custom culture to the workbook
            workbook.Settings.CultureInfo = isoCulture;

            // Save the workbook as JSON; dates will be formatted according to the ISO culture
            workbook.Save(outputPath, SaveFormat.Json);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
