using System;
using Aspose.Cells;

class RemoveThreadedComments
{
    static void Main()
    {
        // Load the workbook (XLSX format)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Determine the used range of the worksheet
            int maxRow = sheet.Cells.MaxDataRow;
            int maxCol = sheet.Cells.MaxDataColumn;

            // Loop through each cell within the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    // Retrieve threaded comments for the current cell
                    ThreadedCommentCollection threadedComments = sheet.Comments.GetThreadedComments(row, col);

                    // If there are any threaded comments, clear them
                    if (threadedComments != null && threadedComments.Count > 0)
                    {
                        threadedComments.Clear();
                    }
                }
            }
        }

        // Save the workbook after removing all threaded comments
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}