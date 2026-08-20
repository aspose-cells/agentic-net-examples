// Title: C# – Merge Multiple JSON Files into a Single Excel Workbook with Separate Worksheets using Aspose.Cells
// Description: A concise C# example that loads a list of JSON files into temporary workbooks with JsonLoadOptions.MultipleWorksheets, combines them into one destination workbook, and saves the result as an XLSX file. The code handles missing files and load errors, making it ideal for consolidating JSON reports, API responses, or log data into distinct Excel sheets.
// Keywords: Aspose.Cells | C# | .NET | JSON to Excel | merge JSON files | multiple worksheets | JsonLoadOptions | Combine workbooks | SaveFormat.Xlsx | data consolidation | Excel automation | API response export | log file analysis
// Common Searches: merge several JSON files into one Excel workbook Aspose.Cells | Aspose.Cells JsonLoadOptions MultipleWorksheets C# example | combine JSON data into separate Excel sheets using .NET | how to load multiple JSON files as worksheets with Aspose.Cells | C# code to consolidate JSON reports into a single XLSX file
// Developer Intent: Create one Excel workbook where each provided JSON file appears on its own worksheet.
// Use Cases: Aggregate daily JSON reports into a master workbook for cross‑day comparison. | Export multiple API response payloads, each on a separate sheet, for stakeholder review. | Combine JSON‑formatted log files into a single Excel file to simplify auditing and filtering.
// AI Prompts: Generate C# code that reads an array of JSON file paths and merges them into a single workbook, placing each top‑level array on a separate worksheet with Aspose.Cells. | Explain how JsonLoadOptions.MultipleWorksheets influences worksheet creation when loading JSON files in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsJsonMergeDemo
{
    // A concise C# example that loads a list of JSON files into temporary workbooks with JsonLoadOptions.MultipleWorksheets, combines them into one destination workbook, and saves the result as an XLSX file. The code handles missing files and load errors, making it ideal for consolidating JSON reports, API responses, or log data into distinct Excel sheets.
    class Program
    {
        static void Main()
        {
            // Paths of the JSON files to be merged
            string[] jsonFiles = new string[]
            {
                "data1.json",
                "data2.json",
                "data3.json"
            };

            // Destination workbook that will contain all worksheets
            Workbook mergedWorkbook = new Workbook();

            // Load each JSON file into a temporary workbook and combine it
            foreach (string jsonPath in jsonFiles)
            {
                if (!File.Exists(jsonPath))
                {
                    Console.WriteLine($"File not found: {jsonPath}. Skipping.");
                    continue;
                }

                try
                {
                    // Load options – set MultipleWorksheets to true if the JSON contains
                    // multiple top‑level array nodes and you want each array on its own sheet.
                    JsonLoadOptions loadOptions = new JsonLoadOptions
                    {
                        MultipleWorksheets = true
                    };

                    // Load the JSON file into a workbook using the options
                    Workbook tempWorkbook = new Workbook(jsonPath, loadOptions);

                    // Combine the temporary workbook into the destination workbook
                    mergedWorkbook.Combine(tempWorkbook);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{jsonPath}': {ex.Message}");
                }
            }

            try
            {
                // Save the merged workbook to an XLSX file
                mergedWorkbook.Save("MergedJsonWorkbook.xlsx", SaveFormat.Xlsx);
                Console.WriteLine("All JSON files have been merged into 'MergedJsonWorkbook.xlsx'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save merged workbook: {ex.Message}");
            }
        }
    }
}
