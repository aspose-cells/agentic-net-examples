using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsJsonMergeDemo
{
    class Program
    {
        static void Main()
        {
            // Paths to the JSON files to be merged
            string[] jsonFiles = new string[]
            {
                "data1.json",
                "data2.json",
                "data3.json"
            };

            // Destination workbook that will contain all worksheets
            Workbook mergedWorkbook = new Workbook();

            // Iterate over each JSON file, load it into a temporary workbook,
            // and combine it with the destination workbook.
            foreach (string jsonPath in jsonFiles)
            {
                if (!File.Exists(jsonPath))
                {
                    Console.WriteLine($"File not found: {jsonPath}");
                    continue;
                }

                // Load options for JSON – we keep default behavior (single worksheet per file)
                JsonLoadOptions loadOptions = new JsonLoadOptions
                {
                    // If the JSON structure contains multiple top‑level arrays and you want each
                    // array as a separate worksheet, set this to true. Here we keep it false
                    // so each file becomes one worksheet.
                    MultipleWorksheets = false
                };

                // Load the JSON file into a temporary workbook
                Workbook tempWorkbook = new Workbook(jsonPath, loadOptions);

                // Combine the temporary workbook into the merged workbook.
                // This adds all worksheets from tempWorkbook to mergedWorkbook.
                mergedWorkbook.Combine(tempWorkbook);
            }

            // Save the merged workbook to an XLSX file
            string outputPath = "MergedJsonWorkbook.xlsx";
            mergedWorkbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Merged workbook saved to: {outputPath}");
        }
    }
}