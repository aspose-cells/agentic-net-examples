// Title: Insert multiple rows with inherited formatting from an Excel template using Aspose.Cells InsertRows (CopyStyle) in C#
// Common Searches: C# Aspose.Cells insert rows and keep template row style | How to use InsertRows with CopyStyle true to preserve formatting in Aspose.Cells | Load an Excel template and add data rows while inheriting cell styles with Aspose.Cells | Aspose.Cells InsertRows copy style example for dynamic record insertion | Preserve row formatting when inserting rows in a workbook using Aspose.Cells C#
// Tags: InsertRows CopyStyle Aspose.Cells C# | inherit row formatting Excel template Aspose | dynamic row insertion Excel workbook Aspose.Cells | load template workbook Aspose.Cells C# | preserve cell style when adding records Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

namespace Example
{
    // The example loads an existing Excel template (or creates a new workbook if the template is missing), inserts the required number of rows at a specified start index using InsertRows with the CopyStyle flag set to true so the new rows inherit the template row's formatting, populates the rows with data, and saves the workbook as an XLSX file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Define template and output file paths
                string templatePath = "template.xlsx";
                string outputPath = "output.xlsx";

                Workbook workbook;

                // Load existing template if it exists; otherwise create a new workbook
                if (File.Exists(templatePath))
                {
                    workbook = new Workbook(templatePath);
                }
                else
                {
                    workbook = new Workbook(); // creates a default workbook with one worksheet
                }

                var worksheet = workbook.Worksheets[0];

                // Sample data to be added as new records
                var records = new List<string[]>
                {
                    new[] { "Alice", "Johnson", "28" },
                    new[] { "Bob", "Williams", "35" }
                };

                // Row index (zero‑based) where new records should start, e.g., after a header row
                int startRow = 1;

                // Insert the required number of rows and inherit the formatting from the template row
                worksheet.Cells.InsertRows(startRow, records.Count, true); // CopyStyle = true

                // Fill the inserted rows with data
                for (int i = 0; i < records.Count; i++)
                {
                    for (int j = 0; j < records[i].Length; j++)
                    {
                        worksheet.Cells[startRow + i, j].PutValue(records[i][j]);
                    }
                }

                // Save the workbook with the new data
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
