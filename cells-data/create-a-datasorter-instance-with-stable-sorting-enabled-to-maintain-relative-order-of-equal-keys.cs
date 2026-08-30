// Title: How to perform a stable multi‑column sort with Aspose.Cells DataSorter in C# while preserving the original order of duplicate rows
// AI Prompts: Create a DataSorter instance, enable its stable‑sorting mode, set two ascending keys (Category and Value), define the cell area to sort, and call sorter.Sort to keep equal‑key rows in their original sequence. | Refactor existing Aspose.Cells C# code to activate stable sorting for a multi‑column sort, specify the sort range, and save the workbook as an .xlsx file.
// Common Searches: Aspose.Cells how to keep original row order when sorting duplicate values in .NET | C# stable sort Excel data using DataSorter in Aspose.Cells | example of multi‑column stable sort with Aspose.Cells .NET | sort range with stability flag Aspose.Cells DataSorter | preserve order of equal keys in Aspose.Cells sort operation
// Tags: stable sorting DataSorter Aspose.Cells | Aspose.Cells multi‑column sort C# | preserve duplicate key order Excel .NET | DataSorter sort range example | Aspose.Cells stable sort workbook

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The sample creates a workbook, fills it with rows that include duplicate categories, configures the workbook's DataSorter with two ascending keys, enables stable sorting, sorts the defined cell area, and saves the result to StableSortedData.xlsx, ensuring rows with equal keys retain their original relative order.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate sample data (including duplicate keys to demonstrate stability)
                cells["A1"].PutValue("Category");
                cells["B1"].PutValue("Value");
                cells["A2"].PutValue("Fruit");
                cells["B2"].PutValue(10);
                cells["A3"].PutValue("Vegetable");
                cells["B3"].PutValue(20);
                cells["A4"].PutValue("Fruit");
                cells["B4"].PutValue(15);
                cells["A5"].PutValue("Vegetable");
                cells["B5"].PutValue(5);
                cells["A6"].PutValue("Fruit");
                cells["B6"].PutValue(12);

                // Get the workbook's DataSorter instance
                DataSorter sorter = workbook.DataSorter;

                // Set the first sort key (Column A) with ascending order
                sorter.Key1 = 0;               // Column A (Category)
                sorter.Order1 = SortOrder.Ascending;

                // Set the second sort key (Column B) with ascending order
                sorter.Key2 = 1;               // Column B (Value)
                sorter.Order2 = SortOrder.Ascending;

                // Define the range to sort (including headers)
                CellArea sortArea = new CellArea
                {
                    StartRow = 0,
                    StartColumn = 0,
                    EndRow = 5,
                    EndColumn = 1
                };

                // Perform the sort
                sorter.Sort(cells, sortArea);

                // Determine output path and ensure directory exists
                string outputPath = "StableSortedData.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to verify the result
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
