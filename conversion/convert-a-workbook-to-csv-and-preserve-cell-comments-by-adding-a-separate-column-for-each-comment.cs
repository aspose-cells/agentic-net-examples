// Title: Convert an Excel workbook to CSV and preserve cell comments by adding comment columns with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export workbook to CSV with comments in separate columns | how to keep Excel cell comments when converting to CSV using Aspose.Cells | C# extract comments from Excel and add them as new columns before CSV export | save each worksheet as individual CSV files while retaining comments Aspose.Cells | insert comment columns programmatically in Aspose.Cells before saving as CSV
// Tags: Aspose.Cells export to CSV with comment columns | C# extract Excel comments into new columns | insert comment column programmatically Aspose.Cells | save each sheet as individual CSV Aspose.Cells | preserve cell comments during CSV conversion .NET

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The example loads an Excel workbook, detects cells that contain comments, inserts a new column after each original column with comments, copies the comment text into the new column, and then saves each worksheet as a separate CSV file using Aspose.Cells' SaveFormat.Csv.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Process each worksheet to extract comments into new columns
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Collect column indexes that contain at least one comment
                List<int> commentColumns = new List<int>();

                foreach (Cell cell in sheet.Cells)
                {
                    if (cell.Comment != null)
                    {
                        int col = cell.Column;
                        if (!commentColumns.Contains(col))
                            commentColumns.Add(col);
                    }
                }

                // Insert new columns after the original ones, starting from the highest index
                commentColumns.Sort((a, b) => b.CompareTo(a));

                foreach (int col in commentColumns)
                {
                    // Insert a new column right after the original column
                    sheet.Cells.InsertColumn(col + 1);

                    // Optional header for the comment column
                    sheet.Cells[0, col + 1].PutValue("Comment");

                    // Determine the last row with data
                    int maxRow = sheet.Cells.MaxDataRow;
                    if (maxRow < 0) maxRow = 0; // handle empty sheets

                    // Transfer comment text to the newly inserted column
                    for (int row = 1; row <= maxRow; row++)
                    {
                        Cell originalCell = sheet.Cells[row, col];
                        if (originalCell.Comment != null)
                        {
                            string commentText = originalCell.Comment.Note;
                            sheet.Cells[row, col + 1].PutValue(commentText);
                        }
                    }
                }
            }

            // Save each worksheet as an individual CSV file
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Worksheet sheet = workbook.Worksheets[i];

                // Create a temporary workbook containing only the current sheet
                Workbook csvWorkbook = new Workbook();

                // Add a copy of the current sheet by name
                csvWorkbook.Worksheets.AddCopy(sheet.Name);

                // Remove the default empty sheet created by the Workbook constructor
                if (csvWorkbook.Worksheets.Count > 1)
                {
                    csvWorkbook.Worksheets.RemoveAt(0);
                }

                // Define CSV file name (e.g., Sheet1.csv, Sheet2.csv, ...)
                string csvFileName = $"Sheet{i + 1}.csv";

                // Save the temporary workbook as CSV
                csvWorkbook.Save(csvFileName, SaveFormat.Csv);
                Console.WriteLine($"Saved {csvFileName}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
