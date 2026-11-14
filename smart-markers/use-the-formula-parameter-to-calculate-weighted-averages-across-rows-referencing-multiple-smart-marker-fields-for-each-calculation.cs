// Title: Calculate per‑row weighted averages in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: c# Aspose.Cells calculate weighted average for each row in Excel | how to export a list of objects to XLSX and compute row totals with Aspose.Cells | using Aspose.Cells to write scores and weights then derive weighted score in C# | generate Excel report with dynamic data and per‑row calculations via Aspose.Cells | Aspose.Cells example for inserting computed values into worksheet from C# collection
// Tags: per‑row score aggregation Aspose.Cells C# | populate Excel worksheet from dynamic collection Aspose.Cells | Aspose.Cells workbook creation and data insertion C# | Excel formula generation with smart markers Aspose.Cells | export dynamic objects to XLSX using Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace WeightedAverageExample
{
    // The program loads an existing template or creates a new workbook, adds column headers, iterates over a list of dynamic objects containing two scores and their corresponding weights, writes each value to successive rows, computes a weighted average for each row using C# arithmetic, writes the result to the worksheet, and saves the workbook as an XLSX file.
    class Program
    {
        static void Main()
        {
            try
            {
                const string templatePath = "Template.xlsx";
                const string resultPath = "Result.xlsx";

                // Load the template workbook if it exists; otherwise create a simple workbook.
                Workbook workbook;
                if (File.Exists(templatePath))
                {
                    workbook = new Workbook(templatePath);
                }
                else
                {
                    workbook = new Workbook();
                    Worksheet ws = workbook.Worksheets[0];

                    // Write headers.
                    ws.Cells["A1"].PutValue("Score1");
                    ws.Cells["B1"].PutValue("Score2");
                    ws.Cells["C1"].PutValue("Weight1");
                    ws.Cells["D1"].PutValue("Weight2");
                    ws.Cells["E1"].PutValue("WeightedAvg");
                }

                // Prepare the data source. Each object represents a row with score and weight fields.
                var rows = new List<dynamic>
                {
                    new { Score1 = 80, Score2 = 90, Weight1 = 0.4, Weight2 = 0.6 },
                    new { Score1 = 70, Score2 = 85, Weight1 = 0.5, Weight2 = 0.5 },
                    new { Score1 = 92, Score2 = 78, Weight1 = 0.3, Weight2 = 0.7 }
                    // Add more rows as needed.
                };

                // Write data and calculate weighted average manually.
                Worksheet worksheet = workbook.Worksheets[0];
                int startRow = 1; // Zero‑based index; row 1 is the second row (below headers).

                foreach (var row in rows)
                {
                    worksheet.Cells[startRow, 0].PutValue(row.Score1);
                    worksheet.Cells[startRow, 1].PutValue(row.Score2);
                    worksheet.Cells[startRow, 2].PutValue(row.Weight1);
                    worksheet.Cells[startRow, 3].PutValue(row.Weight2);

                    // Calculate weighted average.
                    double weightedAvg = (row.Score1 * row.Weight1 + row.Score2 * row.Weight2) /
                                         (row.Weight1 + row.Weight2);
                    worksheet.Cells[startRow, 4].PutValue(weightedAvg);

                    startRow++;
                }

                // Save the resulting workbook.
                workbook.Save(resultPath);
                Console.WriteLine($"Workbook saved successfully to '{resultPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
