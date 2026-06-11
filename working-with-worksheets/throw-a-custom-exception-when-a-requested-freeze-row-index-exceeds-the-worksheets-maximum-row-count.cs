using System;
using Aspose.Cells;

namespace FreezeRowValidationDemo
{
    // Custom exception for invalid freeze row index
    public class FreezeRowIndexOutOfRangeException : Exception
    {
        public FreezeRowIndexOutOfRangeException(string message) : base(message) { }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Desired freeze row index (zero‑based). Change this value as needed.
            int requestedFreezeRow = 70000; // Example that exceeds the limit for older formats

            // Retrieve the maximum row index allowed by the workbook's format
            int maxRowIndex = workbook.Settings.MaxRow;

            // Validate the requested freeze row index against the maximum
            if (requestedFreezeRow > maxRowIndex)
            {
                throw new FreezeRowIndexOutOfRangeException(
                    $"Requested freeze row index {requestedFreezeRow} exceeds the maximum row index {maxRowIndex} for this workbook.");
            }

            // Freeze panes if validation succeeds (freeze rows up to the requested index)
            worksheet.FreezePanes(requestedFreezeRow, 0, requestedFreezeRow, 0);

            // Save the workbook
            workbook.Save("FreezeRowValidated.xlsx");
        }
    }
}