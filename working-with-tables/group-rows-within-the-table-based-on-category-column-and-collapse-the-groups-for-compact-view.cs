using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class GroupRowsByCategory
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Header
                cells["A1"].PutValue("Category");
                cells["B1"].PutValue("Value");

                // Sample data
                string[] categories = { "Alpha", "Alpha", "Alpha", "Beta", "Beta", "Gamma", "Gamma", "Gamma", "Gamma" };
                int[] values = { 10, 15, 20, 30, 35, 40, 45, 50, 55 };

                for (int i = 0; i < categories.Length; i++)
                {
                    int rowIndex = i + 1; // +1 because row 0 is header
                    cells[rowIndex, 0].PutValue(categories[i]); // Column A
                    cells[rowIndex, 1].PutValue(values[i]);    // Column B
                }

                // Group rows by consecutive identical categories
                int startRow = 1; // first data row (zero‑based index)
                string currentCategory = cells[startRow, 0].StringValue;

                for (int row = startRow + 1; row <= categories.Length; row++)
                {
                    bool isEnd = row == categories.Length;
                    string nextCategory = isEnd ? null : cells[row, 0].StringValue;

                    if (isEnd || nextCategory != currentCategory)
                    {
                        int endRow = row - 1; // last row of the current group

                        // Group only if the group has more than one row
                        if (endRow > startRow)
                        {
                            cells.GroupRows(startRow, endRow, true);
                        }

                        // Prepare for the next group
                        if (!isEnd)
                        {
                            startRow = row;
                            currentCategory = nextCategory;
                        }
                    }
                }

                // Show summary row above details (compact view)
                worksheet.Outline.SummaryRowBelow = false;

                // Save the workbook
                string outputPath = "GroupedByCategory.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            GroupRowsByCategory.Run();
        }
    }
}