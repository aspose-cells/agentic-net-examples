// Title: Update Excel Subject Property from Worksheet Content with Aspose.Cells for .NET (C#)
// Description: C# example that creates or loads a workbook, scans every non‑empty cell, determines the most frequent word (case‑insensitive) as the main topic, assigns it to BuiltInDocumentProperties.Subject (and optionally Title), saves the file, reloads it, and verifies that the metadata is persisted.
// Keywords: Aspose.Cells | C# Excel metadata | set Subject property | built‑in document properties | Excel content analysis | extract most frequent word | auto‑generate workbook metadata | save and reload Excel file | global Excel automation | US developers
// Common Searches: How to set Excel Subject property from cell values using Aspose.Cells | C# code to extract main topic of a worksheet and update document metadata | Aspose.Cells example for analyzing worksheet text and setting built‑in properties | Automatically generate Excel metadata based on worksheet content | Determine most common word in Excel sheet with Aspose.Cells
// Developer Intent: Automatically assign the workbook's Subject built‑in property to the most frequent word found in its worksheet cells.
// Use Cases: Generate a descriptive Subject for sales or financial reports without manual entry. | Create a reusable routine that enriches Excel files with meaningful metadata after data population. | Validate that document properties survive the save‑load cycle for compliance or archival purposes.
// AI Prompts: Write a method that scans all non‑empty cells in an Aspose.Cells workbook, returns the most common word, and sets BuiltInDocumentProperties.Subject. | Add robust error handling for cases where the worksheet contains no textual data or only numeric values. | Modify the heuristic to prioritize multi‑word phrases (e.g., "Sales Report") over single words when determining the Subject.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsExamples
{
    // C# example that creates or loads a workbook, scans every non‑empty cell, determines the most frequent word (case‑insensitive) as the main topic, assigns it to BuiltInDocumentProperties.Subject (and optionally Title), saves the file, reloads it, and verifies that the metadata is persisted.
    public class UpdateSubjectBasedOnContentDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate the worksheet with sample data (in a real scenario the workbook would already contain data)
                sheet.Cells["A1"].PutValue("Sales Report");
                sheet.Cells["A2"].PutValue("January");
                sheet.Cells["B2"].PutValue(1200);
                sheet.Cells["A3"].PutValue("February");
                sheet.Cells["B3"].PutValue(1500);
                sheet.Cells["A4"].PutValue("March");
                sheet.Cells["B4"].PutValue(1300);
                sheet.Cells["C1"].PutValue("Report Summary");
                sheet.Cells["C2"].PutValue("Total sales increased compared to previous month.");

                // Analyze the worksheet content to determine the main topic
                // Simple heuristic: find the most frequent word (case‑insensitive) in all non‑empty cells
                Dictionary<string, int> wordCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                int maxRow = sheet.Cells.MaxDataRow;
                int maxCol = sheet.Cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        var cell = sheet.Cells[row, col];
                        if (cell.Value == null) continue;

                        // Split cell text into words using whitespace and punctuation as delimiters
                        string[] words = cell.StringValue
                                             .Split(new char[] { ' ', '\t', '\r', '\n', ',', '.', ';', ':', '!', '?' },
                                                    StringSplitOptions.RemoveEmptyEntries);

                        foreach (string word in words)
                        {
                            if (string.IsNullOrWhiteSpace(word)) continue;

                            if (wordCounts.ContainsKey(word))
                                wordCounts[word]++;
                            else
                                wordCounts[word] = 1;
                        }
                    }
                }

                // Determine the most frequent word; fallback to a default if no words are found
                string mainTopic = "Untitled Document";
                if (wordCounts.Count > 0)
                {
                    mainTopic = wordCounts.OrderByDescending(kv => kv.Value).First().Key;
                }

                // Update the built‑in Subject property with the identified main topic
                workbook.BuiltInDocumentProperties.Subject = mainTopic;

                // Optionally, also set the Title property for completeness
                workbook.BuiltInDocumentProperties.Title = "Generated Report";

                // Save the workbook to verify the property is stored
                string outputPath = "DocumentWithSubject.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                // Load the saved workbook to demonstrate that the Subject property was persisted
                if (File.Exists(outputPath))
                {
                    Workbook loadedWorkbook = new Workbook(outputPath);
                    Console.WriteLine("Subject property set to: " + loadedWorkbook.BuiltInDocumentProperties.Subject);
                    Console.WriteLine("Title property set to: " + loadedWorkbook.BuiltInDocumentProperties.Title);
                }
                else
                {
                    Console.WriteLine("Failed to save the workbook. File not found: " + outputPath);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            UpdateSubjectBasedOnContentDemo.Run();
        }
    }
}
