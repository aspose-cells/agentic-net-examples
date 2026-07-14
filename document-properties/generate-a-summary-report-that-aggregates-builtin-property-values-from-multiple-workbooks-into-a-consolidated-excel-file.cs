using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSummaryReport
{
    class Program
    {
        static void Main()
        {
            try
            {
                // List of source workbook file paths to aggregate
                List<string> sourceFiles = new List<string>
                {
                    "Workbook1.xlsx",
                    "Workbook2.xlsx",
                    "Workbook3.xlsx"
                    // Add more file paths as needed
                };

                // Validate that at least one source file is provided
                if (sourceFiles.Count == 0)
                {
                    Console.WriteLine("No source files provided.");
                    return;
                }

                // Verify the first workbook exists to obtain built‑in property names
                if (!File.Exists(sourceFiles[0]))
                {
                    Console.WriteLine($"File not found: {sourceFiles[0]}");
                    return;
                }

                // Load the first workbook to retrieve the list of built‑in property names
                List<string> propertyNames = new List<string>();
                using (Workbook firstWorkbook = new Workbook(sourceFiles[0]))
                {
                    foreach (var prop in firstWorkbook.BuiltInDocumentProperties)
                    {
                        propertyNames.Add(prop.Name);
                    }
                }

                // Create a new workbook that will hold the consolidated summary
                Workbook summaryWorkbook = new Workbook(); // uses the provided constructor rule
                Worksheet summarySheet = summaryWorkbook.Worksheets[0];

                // Write header row: FileName + each property name
                int headerRow = 0;
                summarySheet.Cells[headerRow, 0].PutValue("FileName");
                for (int i = 0; i < propertyNames.Count; i++)
                {
                    summarySheet.Cells[headerRow, i + 1].PutValue(propertyNames[i]);
                }

                // Iterate over each source workbook and copy property values
                int currentRow = 1;
                foreach (string filePath in sourceFiles)
                {
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"Skipping missing file: {filePath}");
                        continue;
                    }

                    using (Workbook srcWorkbook = new Workbook(filePath))
                    {
                        // Write the file name
                        summarySheet.Cells[currentRow, 0].PutValue(Path.GetFileName(filePath));

                        // Write each built‑in property value
                        for (int i = 0; i < propertyNames.Count; i++)
                        {
                            string propName = propertyNames[i];
                            // Retrieve the property; if missing, write empty string
                            object value = srcWorkbook.BuiltInDocumentProperties[propName]?.Value ?? string.Empty;
                            summarySheet.Cells[currentRow, i + 1].PutValue(value);
                        }
                    }

                    currentRow++;
                }

                // Save the consolidated summary workbook
                string outputPath = "ConsolidatedSummary.xlsx";
                summaryWorkbook.Save(outputPath); // save rule

                Console.WriteLine($"Summary report saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}