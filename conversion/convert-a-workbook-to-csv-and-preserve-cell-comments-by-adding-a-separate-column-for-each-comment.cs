// Title: Convert an Excel workbook to CSV and keep cell comments by inserting separate comment columns with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export Excel to CSV with comments preserved | how to add a comment column when converting .xlsx to .csv using Aspose.Cells | C# insert column after each column that has a comment before saving as CSV | save each worksheet as a separate CSV file while retaining cell notes Aspose.Cells | preserve Excel cell notes during CSV conversion using Aspose.Cells .NET
// Tags: export workbook to CSV with comment columns Aspose.Cells | insert column for cell comments C# Aspose.Cells | preserve Excel notes during CSV conversion .NET | save worksheets as individual CSV files Aspose.Cells | group comments by column before CSV export

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Cells;

namespace WorkbookToCsvWithComments
{
    // The example loads an Excel workbook, iterates through its worksheets, groups cell comments by column, inserts a new column after each column that contains comments, copies the comment text into the new column with an appropriate header, sets the processed sheet as active, and saves each worksheet as a separate CSV file using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Input Excel file path
            string inputPath = @"C:\Input\Sample.xlsx";

            // Output folder for CSV files
            string outputFolder = @"C:\Output\CsvFiles";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Ensure output directory exists
                Directory.CreateDirectory(outputFolder);

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Process each worksheet in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    try
                    {
                        // Get all comments in the current worksheet
                        CommentCollection comments = sheet.Comments;

                        // Group comments by their original column index
                        var commentsByColumn = comments
                            .Cast<Comment>()
                            .GroupBy(c => c.Column)               // use Comment.Column
                            .OrderByDescending(g => g.Key)        // Process from rightmost column to avoid index shift
                            .ToList();

                        // Insert a new column after each column that contains comments
                        foreach (var columnGroup in commentsByColumn)
                        {
                            int originalColumnIndex = columnGroup.Key;
                            int commentColumnIndex = originalColumnIndex + 1; // Insert after the original column

                            // Insert a new column (do not copy style)
                            sheet.Cells.InsertColumn(commentColumnIndex, false);

                            // Set header for the comment column (optional)
                            string originalHeader = sheet.Cells[0, originalColumnIndex].StringValue;
                            sheet.Cells[0, commentColumnIndex].PutValue(
                                string.IsNullOrEmpty(originalHeader)
                                    ? $"Comment_{originalColumnIndex}"
                                    : $"{originalHeader}_Comment");

                            // Populate comment texts in the new column
                            foreach (Comment comment in columnGroup)
                            {
                                int rowIndex = comment.Row; // use Comment.Row
                                string commentText = comment.Note;
                                sheet.Cells[rowIndex, commentColumnIndex].PutValue(commentText);
                            }
                        }

                        // Set the active sheet to the current worksheet before saving as CSV
                        workbook.Worksheets.ActiveSheetIndex = sheet.Index;

                        // Build CSV file name based on worksheet name
                        string csvFileName = Path.Combine(outputFolder, $"{sheet.Name}.csv");

                        // Save the current worksheet as a CSV file
                        workbook.Save(csvFileName, SaveFormat.Csv);
                    }
                    catch (Exception exSheet)
                    {
                        Console.WriteLine($"Error processing sheet '{sheet.Name}': {exSheet.Message}");
                    }
                }

                Console.WriteLine("Conversion completed. CSV files are saved in: " + outputFolder);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
