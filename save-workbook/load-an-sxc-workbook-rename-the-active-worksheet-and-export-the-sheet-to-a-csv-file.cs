// Title: Rename the active worksheet of an SXC workbook and save it as CSV using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells load .sxc file and rename active sheet before exporting to CSV | How to change worksheet name in an OpenDocument spreadsheet with Aspose.Cells .NET | Export a specific worksheet from an SXC workbook to CSV using Aspose.Cells C# | Aspose.Cells .NET example for renaming active worksheet and saving as CSV | Convert renamed sheet of an OpenDocument spreadsheet to CSV in C#
// Tags: load SXC workbook Aspose.Cells .NET | rename active worksheet Aspose.Cells | export worksheet to CSV Aspose.Cells | OpenDocument spreadsheet to CSV conversion | Aspose.Cells SaveFormat.Csv usage

using System;
using Aspose.Cells;

// This example loads an OpenDocument Spreadsheet (SXC) file, retrieves the active worksheet, renames it, and then saves the workbook as a CSV file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Path to the source SXC workbook
        string inputPath = "input.sxc";

        // Load the workbook from the SXC file
        Workbook workbook = new Workbook(inputPath);

        // Get the active worksheet (the one currently selected when the workbook was opened)
        Worksheet activeSheet = workbook.Worksheets[workbook.Worksheets.ActiveSheetIndex];

        // Rename the active worksheet
        activeSheet.Name = "RenamedSheet";

        // Export the renamed worksheet to a CSV file
        // The SaveOptions specify CSV format; the first parameter is the output file name,
        // the second parameter is the index of the worksheet to export.
        workbook.Save("output.csv", SaveFormat.Csv);
    }
}
