using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class UpdateSubjectBasedOnContentDemo
    {
        // Entry point required by the runtime
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (could be replaced with actual data loading)
            sheet.Cells["A1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Marketing");
            sheet.Cells["A3"].PutValue("Sales");
            sheet.Cells["A4"].PutValue("Finance");
            sheet.Cells["A5"].PutValue("Sales");

            // Analyze worksheet content to find the most frequent non‑empty string
            var frequency = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            int maxRow = sheet.Cells.MaxDataRow;
            for (int row = 0; row <= maxRow; row++)
            {
                var cell = sheet.Cells[row, 0]; // column A
                if (cell.Type == CellValueType.IsString)
                {
                    string text = cell.StringValue.Trim();
                    if (!string.IsNullOrEmpty(text))
                    {
                        if (frequency.ContainsKey(text))
                            frequency[text]++;
                        else
                            frequency[text] = 1;
                    }
                }
            }

            // Determine the main topic (the string with highest occurrence)
            string mainTopic = "Untitled";
            if (frequency.Count > 0)
            {
                mainTopic = frequency.OrderByDescending(kv => kv.Value)
                                     .First()
                                     .Key;
            }

            // Update the built‑in Subject property with the identified main topic
            workbook.BuiltInDocumentProperties.Subject = mainTopic;

            // Save the workbook
            string outputPath = "SubjectUpdated.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            // Verify the file exists before loading
            if (File.Exists(outputPath))
            {
                Workbook loaded = new Workbook(outputPath);
                Console.WriteLine("Subject property set to: " + loaded.BuiltInDocumentProperties.Subject);
            }
            else
            {
                Console.WriteLine($"Failed to create output file: {outputPath}");
            }
        }
    }
}