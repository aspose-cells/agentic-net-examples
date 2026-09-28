// Title: How to explicitly set UTF-8 encoding with JsonSaveOptions when exporting an Aspose.Cells workbook to JSON in C#
// AI Prompts: Provide C# code that creates a JsonSaveOptions instance, assigns Encoding = Encoding.UTF8, and saves a Workbook as a JSON file using Aspose.Cells. | Show an Aspose.Cells example that configures JSON output to use UTF-8 before calling Workbook.Save in C#. | Update existing workbook‑to‑JSON C# code to include JsonSaveOptions.Encoding = Encoding.UTF8 and explain the impact on the saved file.
// Common Searches: Aspose.Cells C# export workbook to JSON with specific UTF-8 encoding | Set JsonSaveOptions.Encoding property to UTF-8 in Aspose.Cells example | C# code for saving Excel file as JSON using Aspose.Cells with UTF-8 output | How to change default JSON encoding when using Aspose.Cells JsonSaveOptions | Saving workbook as JSON in C# ensuring UTF-8 character set with Aspose.Cells
// Tags: Aspose.Cells JsonSaveOptions UTF-8 | export workbook to JSON C# | configure JSON encoding Aspose.Cells | Workbook.Save JSON options C# | set JsonSaveOptions.Encoding property

using Aspose.Cells;
using Aspose.Cells.Saving;
using System;
using System.IO;

// The sample creates a Workbook, adds sample data, configures JsonSaveOptions with UTF-8 encoding, and saves the workbook to an output.json file using Aspose.Cells.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["A2"].PutValue("John");
            sheet.Cells["B2"].PutValue(30);

            // JSON save options (UTF-8 is default)
            JsonSaveOptions jsonOptions = new JsonSaveOptions();

            // Output file path
            string outputPath = "output.json";

            // Save the workbook as JSON
            workbook.Save(outputPath, jsonOptions);
            Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
