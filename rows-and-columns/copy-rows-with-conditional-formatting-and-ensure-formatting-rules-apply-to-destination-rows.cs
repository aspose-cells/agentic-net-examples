// Title: Copy worksheet rows and preserve conditional formatting between Excel workbooks using Aspose.Cells for .NET (C#)
// AI Prompts: Copy all used rows from a source worksheet to a new workbook while retaining the original conditional formatting rules with Aspose.Cells in C#. | Transfer the ConditionalFormattings collection from one worksheet to another after copying rows using the Aspose.Cells API. | Duplicate an entire sheet's row data and its conditional formatting into a fresh workbook programmatically with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# copy rows from one Excel file to another preserving conditional formatting | How to duplicate rows and conditional formatting between workbooks using Aspose.Cells | Copy entire worksheet rows with conditional formatting in .NET | Transfer conditional formatting rules after copying rows with Aspose.Cells API
// Tags: copy rows Aspose.Cells C# | conditional formatting transfer Aspose.Cells | duplicate worksheet rows .NET | preserve Excel conditional formatting programmatically | copy rows between workbooks Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsCopyRowsWithConditionalFormatting
{
    // Loads a source Excel file, copies all used rows from its first worksheet to a new workbook, copies the worksheet's ConditionalFormattings collection, and saves the result as destination.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for source and destination workbooks
            string sourcePath = "source.xlsx";
            string destinationPath = "destination.xlsx";

            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(sourcePath);
            // Create a new workbook for the destination
            Workbook destinationWorkbook = new Workbook();

            // Access the first worksheet in both workbooks (adjust index if needed)
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            Worksheet destSheet = destinationWorkbook.Worksheets[0];

            // Define the rows to copy
            int sourceRowIndex = 0;          // first row to copy (zero‑based)
            int destinationRowIndex = 0;    // where to start pasting in destination
            int rowCount = sourceSheet.Cells.MaxDisplayRange.RowCount; // copy all used rows

            // Copy rows including data and cell formats
            destSheet.Cells.CopyRows(sourceSheet.Cells, sourceRowIndex, destinationRowIndex, rowCount);

            // Copy conditional formatting rules from source to destination
            destSheet.ConditionalFormattings.Copy(sourceSheet.ConditionalFormattings);

            // Save the result
            destinationWorkbook.Save(destinationPath);
        }
    }
}
