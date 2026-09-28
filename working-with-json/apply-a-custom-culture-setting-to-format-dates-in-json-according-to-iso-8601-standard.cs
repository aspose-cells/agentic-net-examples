// Title: Use an invariant culture to ensure ISO 8601 date formatting when saving an Aspose.Cells workbook to JSON in C#
// AI Prompts: Write C# code that clones InvariantCulture, assigns it to CultureInfo.CurrentCulture, and then saves a Workbook with JsonSaveOptions so all dates appear as ISO 8601 strings in the JSON file. | Adapt an existing Aspose.Cells JSON export routine to apply a custom culture that forces UTC ISO 8601 formatting for every DateTime cell before calling Workbook.Save.
// Common Searches: how to make Aspose.Cells JSON export use ISO 8601 dates in C# | set CultureInfo for JSON save options in Aspose.Cells | C# export Excel workbook to JSON with invariant culture | ensure UTC date format in JSON output from Aspose.Cells workbook | Aspose.Cells JsonSaveOptions date format customization
// Tags: Aspose.Cells JSON export ISO 8601 dates | C# set invariant culture for Aspose.Cells | JsonSaveOptions date formatting control | Excel to JSON date format C# | CultureInfo.CurrentCulture custom culture Aspose.Cells

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// The program clones the invariant culture, sets it as the current thread culture, creates a workbook with DateTime values, and saves the workbook as JSON using Aspose.Cells, which outputs all dates in ISO 8601 format.
class Program
{
    static void Main()
    {
        try
        {
            // Set a culture that does not affect ISO 8601 formatting.
            CultureInfo isoCulture = (CultureInfo)CultureInfo.InvariantCulture.Clone();
            CultureInfo.CurrentCulture = isoCulture;

            // Create a new workbook and get the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add headers.
            sheet.Cells["A1"].PutValue("Id");
            sheet.Cells["B1"].PutValue("CreatedDate");

            // Add sample data with DateTime values.
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue(DateTime.Now); // Local time
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue(new DateTime(2023, 5, 15, 14, 30, 0, DateTimeKind.Utc)); // UTC time

            // Configure JSON save options (default ISO 8601 format is used).
            JsonSaveOptions jsonOptions = new JsonSaveOptions();

            // Define output file path.
            string outputPath = "output.json";

            // Save the workbook as JSON.
            workbook.Save(outputPath, jsonOptions);
            Console.WriteLine($"Workbook successfully saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
