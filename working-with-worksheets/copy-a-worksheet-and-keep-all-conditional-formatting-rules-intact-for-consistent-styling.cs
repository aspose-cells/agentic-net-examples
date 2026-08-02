using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the source workbook containing the worksheet with conditional formatting
        Workbook sourceWorkbook = new Workbook("source.xlsx");
        Worksheet sourceSheet = sourceWorkbook.Worksheets[0]; // adjust index or name as needed

        // Create a new workbook that will receive the copied worksheet
        Workbook destinationWorkbook = new Workbook();

        // Add a copy of the source worksheet to the destination workbook
        // AddCopy copies contents and formats (including conditional formatting in most cases)
        int copiedIndex = destinationWorkbook.Worksheets.AddCopy(sourceSheet.Name);
        Worksheet destinationSheet = destinationWorkbook.Worksheets[copiedIndex];

        // Explicitly copy conditional formatting to guarantee it is retained
        destinationSheet.ConditionalFormattings.Copy(sourceSheet.ConditionalFormattings);

        // Save the resulting workbook
        destinationWorkbook.Save("destination.xlsx", SaveFormat.Xlsx);
    }
}
// Author: Aspose.Cells .NET example – copies a worksheet while preserving conditional formatting.