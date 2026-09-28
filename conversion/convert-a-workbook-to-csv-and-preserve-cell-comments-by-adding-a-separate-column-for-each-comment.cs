// Title: Convert each Excel worksheet to a separate CSV file while preserving cell comments by adding comment columns with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx workbook using Aspose.Cells, iterates through all worksheets, creates a new column for every cell comment (header shows the original cell address), copies the sheet into a temporary workbook, and saves it as a CSV file. | Adapt an existing Aspose.Cells .NET conversion routine to insert comment columns before exporting so that each generated CSV includes the comment text alongside the original data for every worksheet.
// Common Searches: Aspose.Cells .NET export worksheet to CSV and keep comments | How to add comment columns when converting Excel to CSV using C# | Save each sheet of an Excel file as separate CSV files with cell notes using Aspose.Cells | Preserve Excel cell comments in CSV output with Aspose.Cells for .NET | C# generate CSV per worksheet and include comment text as separate columns
// Tags: worksheet to CSV conversion Aspose.Cells | preserve cell comments Aspose.Cells | insert comment columns Excel C# | export separate CSV per sheet .NET | comment extraction to CSV using Aspose API

using System;
using System.IO;
using Aspose.Cells;

// The program loads input.xlsx, iterates each worksheet, inserts a new column for every cell comment (header contains the original cell address), copies the modified sheet into a temporary workbook, and saves each sheet as an individual CSV file using Aspose.Cells.
class WorkbookToCsvWithComments
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Process each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    // Get all comments in the current worksheet
                    CommentCollection comments = sheet.Comments;
                    int commentCount = comments.Count;

                    if (commentCount > 0)
                    {
                        // Determine the index of the first new column (after the last used column)
                        int firstNewColumnIndex = sheet.Cells.MaxColumn + 1;

                        // Insert enough columns to hold all comments
                        sheet.Cells.InsertColumns(firstNewColumnIndex, commentCount);

                        // Populate each new column with the comment text
                        for (int i = 0; i < commentCount; i++)
                        {
                            Comment comment = comments[i];

                            // Column that will store this comment
                            int commentColumn = firstNewColumnIndex + i;

                            // Header showing the original cell address (e.g., "A1")
                            string cellAddress = CellsHelper.CellIndexToName(comment.Row, comment.Column);
                            sheet.Cells[0, commentColumn].PutValue($"Comment_{cellAddress}");

                            // Place the comment text in the same row as the original cell
                            sheet.Cells[comment.Row, commentColumn].PutValue(comment.Note);
                        }
                    }

                    // Export the processed worksheet to a CSV file (one CSV per worksheet)
                    string csvFileName = $"{sheet.Name}.csv";

                    // Create a temporary workbook containing only the current sheet
                    Workbook tempWb = new Workbook();
                    tempWb.Worksheets.Clear();

                    // Add a copy of the current sheet by name (AddCopy overload expects a string)
                    tempWb.Worksheets.AddCopy(sheet.Name);
                    tempWb.Worksheets[0].PageSetup.PrintArea = ""; // Ensure entire sheet is exported

                    // Save as CSV
                    tempWb.Save(csvFileName, SaveFormat.Csv);
                }
                catch (Exception exSheet)
                {
                    Console.WriteLine($"Error processing sheet \"{sheet.Name}\": {exSheet.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
