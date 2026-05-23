using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace QuarterlyFinancialMapping
{
    class Program
    {
        static void Main()
        {
            try
            {
                // -------------------- Lifecycle: Create Workbook --------------------
                // If you have a template, verify its existence before loading.
                // string templatePath = "Template.xlsx";
                // if (File.Exists(templatePath))
                //     workbook = new Workbook(templatePath);
                // else
                //     workbook = new Workbook(); // create new workbook

                Workbook workbook = new Workbook(); // new workbook
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // -------------------- Prepare Sample Quarterly Data --------------------
                // Each inner array represents Q1‑Q4 results for a specific year
                List<double[]> quarterlyResults = new List<double[]>
                {
                    new double[] { 12500, 13800, 14250, 15000 }, // Year 2022
                    new double[] { 16000, 16500, 17000, 17500 }, // Year 2023
                    new double[] { 18000, 18500, 19000, 19500 }  // Year 2024
                };

                // -------------------- Define Layout in the Worksheet --------------------
                // Assume the financial statement layout starts at cell A2:
                //   A column  -> Year label
                //   B‑E columns -> Q1‑Q4 values
                int startRow = 1;          // zero‑based index (A2)
                int yearColumn = 0;        // column A
                int dataStartColumn = 1;   // column B (first quarter)

                // Write header row (optional, for clarity)
                cells["A1"].PutValue("Year");
                cells["B1"].PutValue("Q1");
                cells["C1"].PutValue("Q2");
                cells["D1"].PutValue("Q3");
                cells["E1"].PutValue("Q4");

                // -------------------- Map Years --------------------
                for (int i = 0; i < quarterlyResults.Count; i++)
                {
                    int year = 2022 + i;
                    cells[startRow + i, yearColumn].PutValue(year);
                }

                // -------------------- Map Quarterly Values Using Range --------------------
                // Build a 2‑dimensional object array compatible with Range.Value
                object[,] dataArray = new object[quarterlyResults.Count, 4];
                for (int i = 0; i < quarterlyResults.Count; i++)
                {
                    for (int q = 0; q < 4; q++)
                    {
                        dataArray[i, q] = quarterlyResults[i][q];
                    }
                }

                // Create a range that exactly covers the target cells for quarterly data
                // Parameters: firstRow, firstColumn, totalRows, totalColumns
                AsposeRange dataRange = cells.CreateRange(startRow, dataStartColumn, quarterlyResults.Count, 4);
                dataRange.Value = dataArray; // assign the whole 2‑D array in one operation

                // -------------------- (Optional) Apply Simple Formatting --------------------
                // Bold the header row
                Style headerStyle = workbook.CreateStyle();
                headerStyle.Font.IsBold = true;
                StyleFlag flag = new StyleFlag { FontBold = true };
                cells.CreateRange("A1", "E1").ApplyStyle(headerStyle, flag);

                // -------------------- Lifecycle: Save Workbook --------------------
                string outputPath = "QuarterlyFinancialStatement.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}