using System;
using Aspose.Cells;

class ExportTimeline
{
    static void Main()
    {
        // Path to the source workbook that contains the Timeline control
        string sourcePath = "SourceWorkbook.xlsx";

        // Path for the new workbook that will receive the exported Timeline
        string destinationPath = "ExportedTimeline.xlsx";

        // Load the source workbook (preserves all data, formatting, and Timeline objects)
        Workbook sourceWorkbook = new Workbook(sourcePath);

        // Create an empty workbook that will receive the copied content
        Workbook destinationWorkbook = new Workbook();

        // Copy the entire source workbook into the destination workbook.
        // This operation copies worksheets, cell data, formatting, and Timeline controls.
        destinationWorkbook.Copy(sourceWorkbook);

        // Save the new workbook as XLSX, keeping all original formatting and Timeline intact
        destinationWorkbook.Save(destinationPath, SaveFormat.Xlsx);
    }
}