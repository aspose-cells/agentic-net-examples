using System;
using Aspose.Cells;

namespace AsposeCellsCopyMoveDemo
{
    public class Program
    {
        public static void Main()
        {
            // -------------------------------------------------
            // 1. Load the source workbook (the workbook that contains data to be copied)
            // -------------------------------------------------
            Workbook sourceWorkbook = new Workbook("Source.xlsx");

            // -------------------------------------------------
            // 2. Create an empty destination workbook (the target workbook)
            // -------------------------------------------------
            Workbook destinationWorkbook = new Workbook();

            // -------------------------------------------------
            // 3. Example A: Copy the entire source workbook into the destination workbook
            //    This copies all worksheets, styles, and other workbook‑level information.
            // -------------------------------------------------
            destinationWorkbook.Copy(sourceWorkbook);

            // -------------------------------------------------
            // 4. Example B: Copy a single worksheet from the source workbook
            //    to a specific worksheet in the destination workbook.
            // -------------------------------------------------
            // Ensure the destination workbook has a worksheet to receive the copy.
            Worksheet targetSheet = destinationWorkbook.Worksheets[0];

            // Configure copy options (optional). Here we keep formulas referring to the destination sheet.
            CopyOptions copyOptions = new CopyOptions();
            copyOptions.ReferToDestinationSheet = true;

            // Copy the first worksheet of the source workbook into the target sheet.
            sourceWorkbook.Worksheets[0].Copy(targetSheet, copyOptions);

            // -------------------------------------------------
            // 5. Move the copied worksheet to a new position within the destination workbook.
            //    For example, move it to index 2 (third position, zero‑based indexing).
            // -------------------------------------------------
            targetSheet.MoveTo(2);

            // -------------------------------------------------
            // 6. Save the destination workbook to disk.
            // -------------------------------------------------
            destinationWorkbook.Save("Result.xlsx", SaveFormat.Xlsx);
        }
    }
}