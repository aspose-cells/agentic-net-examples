// Title: How to enforce ISO‑8601 date format when saving an Aspose.Cells workbook to JSON using InvariantCulture in C#
// Common Searches: c# Aspose.Cells save workbook to JSON with ISO 8601 dates | how to use InvariantCulture with Aspose.Cells JsonSaveOptions | set workbook Settings.CultureInfo before exporting to JSON in .NET | Aspose.Cells date formatting in JSON output using culture settings | export Excel dates as ISO 8601 string with Aspose.Cells C#
// Tags: Aspose.Cells workbook cultureinfo configuration | JSON export date format ISO 8601 Aspose.Cells | C# JsonSaveOptions invariant culture | Excel to JSON conversion with standardized dates | Aspose.Cells date serialization settings

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The example creates a workbook, inserts DateTime values, sets the workbook's Settings.CultureInfo to InvariantCulture to guarantee culture‑independent formatting, configures JsonSaveOptions, and saves the workbook as JSON where dates are automatically rendered in ISO‑8601 format.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Insert sample date values
            sheet.Cells["A1"].PutValue(DateTime.Now);
            sheet.Cells["A2"].PutValue(new DateTime(2023, 5, 15, 13, 45, 30));

            // Apply invariant culture to ensure culture‑independent formatting
            workbook.Settings.CultureInfo = CultureInfo.InvariantCulture;

            // Configure JSON save options (default ISO 8601 format is used)
            JsonSaveOptions jsonOptions = new JsonSaveOptions();

            // Define output file path
            string outputPath = "output.json";

            // Save the workbook as JSON using the specified options
            workbook.Save(outputPath, jsonOptions);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
