// Title: Sort column M by custom priority order (High, Medium, Low) using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells sort Excel column by custom priority list High Medium Low | How to implement custom sort order for a specific column using Aspose.Cells DataSorter | Using a helper column to map text priorities to numbers for sorting in Aspose.Cells | Aspose.Cells example to sort column M while preserving original data in C# | Custom sort order for Excel data with Aspose.Cells .NET library
// Tags: custom priority sorting Aspose.Cells | DataSorter helper column mapping Aspose.Cells | sort specific column M Aspose.Cells C# | Excel custom order High Medium Low Aspose.Cells | temporary numeric column for sorting Aspose.Cells | remove helper column after sort Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace CustomSortExample
{
    // The example loads an XLSX workbook, adds a temporary helper column that converts the text values "High", "Medium", and "Low" in column M to numeric ranks, configures Aspose.Cells' DataSorter to sort by that helper column in ascending order, deletes the helper column after sorting, and saves the sorted workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "sorted.xlsx";

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);
                Worksheet worksheet = workbook.Worksheets[0]; // work with the first worksheet

                // Determine the used range for sorting
                int startRow = 0; // include header row; change to 1 to skip headers
                int endRow = worksheet.Cells.MaxDataRow;
                int startColumn = 0;
                int endColumn = worksheet.Cells.MaxDataColumn;

                // Ensure column M (index 12) is within the used range
                const int priorityColumnIndex = 12; // Column M (zero‑based)
                if (priorityColumnIndex > endColumn)
                {
                    Console.WriteLine("Error: Column M is outside the used range of the worksheet.");
                    return;
                }

                // Add a temporary helper column (next column) to hold numeric priority values
                int helperColumnIndex = endColumn + 1;
                for (int row = startRow; row <= endRow; row++)
                {
                    string priority = worksheet.Cells[row, priorityColumnIndex].StringValue?.Trim();
                    int numericPriority = priority switch
                    {
                        "High" => 1,
                        "Medium" => 2,
                        "Low" => 3,
                        _ => 4 // any other value goes to the end
                    };
                    worksheet.Cells[row, helperColumnIndex].PutValue(numericPriority);
                }

                // Extend the sort area to include the helper column
                CellArea sortArea = new CellArea
                {
                    StartRow = startRow,
                    EndRow = endRow,
                    StartColumn = startColumn,
                    EndColumn = helperColumnIndex
                };

                // Configure the DataSorter to sort by the helper column (ascending)
                DataSorter sorter = workbook.DataSorter;
                sorter.Key1 = helperColumnIndex;
                sorter.Order1 = SortOrder.Ascending;

                // Perform the sort
                sorter.Sort(worksheet.Cells, sortArea);

                // Optionally remove the helper column after sorting
                worksheet.Cells.DeleteColumn(helperColumnIndex);

                // Save the sorted workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Sorting completed successfully. Output saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An unexpected error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
