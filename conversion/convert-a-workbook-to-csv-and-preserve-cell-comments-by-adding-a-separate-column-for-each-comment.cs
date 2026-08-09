// Title: C# – Convert Excel to CSV with Aspose.Cells while preserving cell comments as separate columns
// Description: Loads an Excel workbook, adds a dedicated column for each cell comment (headered with the cell address), copies the comment text into that column, and saves every worksheet as an individual CSV file using Aspose.Cells TxtSaveOptions. The solution works for multiple sheets and creates an output folder for the generated CSVs.
// Keywords: Aspose.Cells CSV export | C# Excel to CSV conversion | preserve Excel comments | comment columns in CSV | Aspose.Cells workbook to CSV | export cell notes to CSV | .NET Excel comment handling
// Common Searches: Aspose.Cells export comments to CSV C# | include Excel cell notes when saving as CSV | add comment column during CSV conversion Aspose | C# convert workbook to CSV with comments | preserve Excel comments in CSV output
// Developer Intent: Export each worksheet of an Excel file to CSV while keeping the original cell comments in separate columns.
// Use Cases: Auditable CSV reports that retain reviewer notes. | Data pipelines that require comment context alongside values. | Batch processing of workbooks where comments must travel with the extracted data.
// AI Prompts: Show a C# Aspose.Cells example that reads an Excel file, creates comment columns, and saves each sheet as CSV. | Explain how to optimize the code for workbooks containing thousands of comments. | Suggest ways to customize comment column headers to include sheet name, comment author, or timestamp.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCommentToCsv
{
    // Loads an Excel workbook, adds a dedicated column for each cell comment (headered with the cell address), copies the comment text into that column, and saves every worksheet as an individual CSV file using Aspose.Cells TxtSaveOptions. The solution works for multiple sheets and creates an output folder for the generated CSVs.
    class Program
    {
        static void Main()
        {
            try
            {
                // Input Excel workbook path
                string inputPath = "input.xlsx";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                    return;
                }

                // Folder to store generated CSV files
                string outputFolder = "CsvOutput";
                Directory.CreateDirectory(outputFolder);

                // Load the workbook (lifecycle rule: use Workbook constructor)
                Workbook workbook = new Workbook(inputPath);

                // Iterate through each worksheet in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Process comments: add a separate column for each comment
                    CommentCollection comments = sheet.Comments;
                    int commentCount = comments.Count;

                    if (commentCount > 0)
                    {
                        // Determine the last used column index in the sheet
                        int lastUsedColumn = sheet.Cells.MaxColumn;

                        // Add a column for each comment
                        for (int i = 0; i < commentCount; i++)
                        {
                            Comment comment = comments[i];

                            // Determine column for this comment (after existing data)
                            int commentColumnIndex = lastUsedColumn + 1 + i;

                            // Header for the comment column (e.g., "Comment_A1")
                            string cellAddress = sheet.Cells[comment.Row, comment.Column].Name;
                            string header = $"Comment_{cellAddress}";
                            sheet.Cells[0, commentColumnIndex].PutValue(header);

                            // Place the comment text in the same row as the commented cell
                            sheet.Cells[comment.Row, commentColumnIndex].PutValue(comment.Note);
                        }
                    }

                    // Create a temporary workbook that contains only the processed sheet
                    Workbook tempWorkbook = new Workbook();
                    tempWorkbook.Worksheets.Clear();

                    // Add a copy of the current sheet to the temporary workbook
                    tempWorkbook.Worksheets.AddCopy(sheet.Name);

                    // Prepare CSV save options (TxtSaveOptions with ExportAllSheets = true)
                    TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
                    {
                        ExportAllSheets = true
                    };

                    // Build CSV file path
                    string csvPath = Path.Combine(outputFolder, $"{sheet.Name}.csv");

                    // Save the temporary workbook as CSV (lifecycle rule: use Save with SaveOptions)
                    tempWorkbook.Save(csvPath, csvOptions);
                }

                Console.WriteLine("Conversion completed. CSV files are located in: " + Path.GetFullPath(outputFolder));
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
