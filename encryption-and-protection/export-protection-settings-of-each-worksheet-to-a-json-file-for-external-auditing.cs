// Title: Export each worksheet's protection status to JSON using Aspose.Cells for .NET
// Common Searches: how to extract worksheet protection flags with Aspose.Cells and write them to a JSON file in C# | C# Aspose.Cells code to audit Excel sheets for protection and password presence | serialize Excel worksheet IsProtected property to JSON for external review using Aspose.Cells | retrieve password‑set indicator for every worksheet in a workbook with Aspose.Cells .NET | list protected worksheets and export their status to a JSON report in C#
// Tags: export worksheet protection to JSON | Aspose.Cells worksheet IsProtected extraction | C# audit Excel sheet protection | serialize Excel protection settings Aspose.Cells | list worksheet password presence .NET

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

// The example loads an Excel workbook with Aspose.Cells, iterates through all worksheets, captures each sheet's name, protection flag, and whether a password is set, then serializes this data to a formatted JSON file for auditing purposes, with error handling for missing files and per‑sheet issues.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "worksheet_protection.json";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            var protectionInfoList = new List<object>();

            // Collect protection details for each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    var protection = sheet.Protection;

                    var sheetInfo = new
                    {
                        SheetName = sheet.Name,
                        IsProtected = sheet.IsProtected,               // correct property on Worksheet
                        HasPassword = !string.IsNullOrEmpty(protection.Password) // password presence flag
                    };

                    protectionInfoList.Add(sheetInfo);
                }
                catch (Exception exSheet)
                {
                    Console.WriteLine($"Error processing sheet '{sheet.Name}': {exSheet.Message}");
                }
            }

            // Serialize the information to formatted JSON
            string json = JsonSerializer.Serialize(protectionInfoList, new JsonSerializerOptions { WriteIndented = true });

            // Write JSON to the output file
            File.WriteAllText(outputPath, json);
            Console.WriteLine($"Worksheet protection information saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
