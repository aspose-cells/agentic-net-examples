using System;
using Aspose.Cells;

namespace WorksheetCopyExample
{
    class Program
    {
        static void Main()
        {
            // Load the source workbook from a file
            Workbook sourceWorkbook = new Workbook("source.xlsx");

            // Create a new (empty) destination workbook
            Workbook destinationWorkbook = new Workbook();
            // Remove the default worksheet that Aspose.Cells creates
            destinationWorkbook.Worksheets.Clear();

            // Configure copy options (optional, keeps formula references within the same workbook)
            CopyOptions copyOptions = new CopyOptions();
            copyOptions.ReferToSheetWithSameName = true;

            // Iterate through each worksheet in the source workbook
            foreach (Worksheet sourceSheet in sourceWorkbook.Worksheets)
            {
                // Add a new worksheet to the destination workbook with the same name
                Worksheet destSheet = destinationWorkbook.Worksheets.Add(sourceSheet.Name);

                // Copy the contents and formats from the source worksheet to the new worksheet
                destSheet.Copy(sourceSheet, copyOptions);
            }

            // Save the destination workbook to a file
            destinationWorkbook.Save("destination.xlsx");
        }
    }
}