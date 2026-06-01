using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class WorkbookToCsvWithComments
{
    static void Main()
    {
        // Path to the source workbook (can be .xlsx, .xls, etc.)
        string sourcePath = "input.xlsx";

        // Path for the resulting CSV file
        string csvPath = "output.csv";

        // Load the workbook
        Workbook workbook = new Workbook(sourcePath);

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Get the collection of comments on the current sheet
            CommentCollection comments = sheet.Comments;

            // If there are no comments, continue to the next sheet
            if (comments.Count == 0)
                continue;

            // Determine the last used column index in the sheet
            // Cells.MaxColumn returns the zero‑based index of the rightmost column that contains data
            int lastDataColumn = sheet.Cells.MaxColumn;

            // Add a separate column for each comment
            for (int i = 0; i < comments.Count; i++)
            {
                Comment comment = comments[i];

                // Calculate the index of the new column for this comment
                int commentColumnIndex = lastDataColumn + i + 1;

                // Optional: write a header indicating which cell the comment belongs to
                // Example header: "Comment_A1"
                string header = $"Comment_{comment.Row + 1}{CellIndexToName(comment.Column)}";
                sheet.Cells[0, commentColumnIndex].PutValue(header);

                // Write the comment text in the same row as the original cell
                sheet.Cells[comment.Row, commentColumnIndex].PutValue(comment.Note);
            }
        }

        // Save the modified workbook as CSV, exporting all sheets into a single CSV file
        TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv);
        saveOptions.ExportAllSheets = true;
        workbook.Save(csvPath, saveOptions);
    }

    // Helper method to convert a zero‑based column index to its Excel column name (A, B, …, AA, AB, …)
    private static string CellIndexToName(int columnIndex)
    {
        const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        string name = string.Empty;
        int dividend = columnIndex + 1;

        while (dividend > 0)
        {
            int modulo = (dividend - 1) % 26;
            name = letters[modulo] + name;
            dividend = (dividend - modulo) / 26;
        }

        return name;
    }
}