// Title: Group rows by a category column and collapse them in an Excel worksheet with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# group rows by column value and hide the group | How to create outline groups in an Excel sheet using Aspose.Cells .NET | Programmatically collapse rows with the same category in Excel using Aspose.Cells | C# code to hide grouped rows based on a category column with Aspose.Cells | Excel row grouping and collapsing example using Aspose.Cells for .NET
// Tags: group rows by column value Aspose.Cells .NET | collapse outline groups Excel C# | hide grouped rows programmatically Aspose.Cells | create row outlines based on category Aspose.Cells | Excel worksheet row grouping Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace Example
{
    // The sample loads an Excel workbook, scans the first worksheet, groups consecutive rows that share the same value in a designated category column, collapses each group by hiding the rows, and saves the modified file.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Ensure the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);
                Worksheet worksheet = workbook.Worksheets[0];

                // Column that contains the category (0 = column A)
                int categoryColumnIndex = 0;

                // Data range (assumes first row is header)
                int firstDataRow = 1; // zero‑based index
                int lastDataRow = worksheet.Cells.MaxDataRow;

                if (lastDataRow < firstDataRow)
                {
                    Console.WriteLine("No data rows found.");
                    return;
                }

                // Track current group
                int groupStartRow = firstDataRow;
                string previousCategory = worksheet.Cells[groupStartRow, categoryColumnIndex]?.StringValue ?? string.Empty;

                // Detect category changes and create outline groups
                for (int row = firstDataRow + 1; row <= lastDataRow; row++)
                {
                    string currentCategory = worksheet.Cells[row, categoryColumnIndex]?.StringValue ?? string.Empty;

                    if (!currentCategory.Equals(previousCategory, StringComparison.Ordinal))
                    {
                        // Close previous group if it has more than one row
                        if (row - 1 > groupStartRow)
                        {
                            // Group rows and hide them
                            int totalRows = (row - 1) - groupStartRow + 1;
                            worksheet.Cells.GroupRows(groupStartRow, totalRows, true);
                            for (int i = groupStartRow; i <= row - 1; i++)
                            {
                                worksheet.Cells.Rows[i].IsHidden = true;
                            }
                        }

                        // Start new group
                        groupStartRow = row;
                        previousCategory = currentCategory;
                    }
                }

                // Handle the final group
                if (lastDataRow > groupStartRow)
                {
                    int totalRows = lastDataRow - groupStartRow + 1;
                    worksheet.Cells.GroupRows(groupStartRow, totalRows, true);
                    for (int i = groupStartRow; i <= lastDataRow; i++)
                    {
                        worksheet.Cells.Rows[i].IsHidden = true;
                    }
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
