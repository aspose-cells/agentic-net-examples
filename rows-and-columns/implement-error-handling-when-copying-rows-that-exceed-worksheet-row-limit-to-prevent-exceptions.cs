// Title: Validate worksheet row limit before copying rows with Aspose.Cells in C#
// Common Searches: how to prevent Aspose.Cells row overflow when copying rows in C# | C# code to check Excel max rows before using CopyRows with Aspose.Cells | adjust number of rows to copy to stay within .xlsx row limit using Aspose.Cells | handle exception for copying rows beyond worksheet row limit in Aspose.Cells | safe row copy pattern with MaxRow validation in Aspose.Cells C#
// Tags: row limit validation Aspose.Cells | copy rows with MaxRow check C# | prevent worksheet row overflow Aspose.Cells | adjustable row copy count .xlsx | exception handling for CopyRows Aspose.Cells

using System;
using Aspose.Cells;

// The example loads a workbook, retrieves the destination worksheet's MaxRow value, verifies that the requested copy range does not exceed the Excel row limit, adjusts the row count if necessary, and then copies rows inside a try‑catch block, providing console warnings or errors and saving the result.
class CopyRowsWithLimitCheck
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("Input.xlsx");

        // Get source and destination worksheets
        Worksheet srcSheet = workbook.Worksheets[0];
        Worksheet destSheet = workbook.Worksheets[1];

        // Define the range of rows to copy (0‑based index)
        int startRow = 100;          // first row to copy
        int totalRowsToCopy = 5000;  // number of rows to copy

        // Define the destination start row
        int destStartRow = 1048000;  // example that may exceed the limit

        // Get the maximum row index allowed in the worksheet (0‑based)
        int maxRowIndex = destSheet.Cells.MaxRow; // for .xlsx this is 1,048,575

        // Calculate the last row index after copying
        long lastRowIndex = (long)destStartRow + totalRowsToCopy - 1;

        // Check if the copy operation would exceed the worksheet row limit
        if (lastRowIndex > maxRowIndex)
        {
            // Adjust the number of rows to copy so it fits within the limit
            int allowableRows = (int)(maxRowIndex - destStartRow + 1);
            if (allowableRows <= 0)
            {
                Console.WriteLine("Error: Destination start row is beyond the worksheet row limit. No rows will be copied.");
                // Optionally, you could abort the operation here.
                return;
            }

            Console.WriteLine($"Warning: Requested copy exceeds row limit. Adjusting rows to copy from {totalRowsToCopy} to {allowableRows}.");
            totalRowsToCopy = allowableRows;
        }

        // Perform the copy row operation within a try‑catch block for safety
        try
        {
            // Copy rows one by one (you could also use CopyRows if the range is contiguous)
            for (int i = 0; i < totalRowsToCopy; i++)
            {
                int srcRowIndex = startRow + i;
                int destRowIndex = destStartRow + i;

                // Copy the entire row
                destSheet.Cells.CopyRows(srcSheet.Cells, srcRowIndex, destRowIndex, 1);
            }

            Console.WriteLine("Rows copied successfully.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected exceptions gracefully
            Console.WriteLine($"Exception occurred during row copy: {ex.Message}");
        }

        // Save the workbook (replace with your desired output path)
        workbook.Save("Output.xlsx");
    }
}
