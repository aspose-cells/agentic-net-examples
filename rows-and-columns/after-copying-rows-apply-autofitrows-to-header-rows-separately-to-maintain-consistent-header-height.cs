// Title: Copy rows between two workbooks and auto‑fit only the header rows using Aspose.Cells for .NET (C#)
// AI Prompts: Copy the first two rows from a source worksheet to a destination worksheet and then call AutoFitRows on rows 0‑1 only using Aspose.Cells in C#. | Transfer all data rows after the header from one workbook to another while preserving the original header height by applying AutoFitRows exclusively to the header rows with Aspose.Cells. | Save the combined workbook to an XLSX file on the desktop after copying rows and selectively auto‑fitting the header rows using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# copy rows from one workbook to another and auto fit header rows | How to apply AutoFitRows to specific rows after copying rows in Aspose.Cells .NET | Preserve header row height when copying rows with Aspose.Cells for .NET | Selective AutoFitRows for header rows in C# Aspose.Cells example
// Tags: row copying with selective AutoFitRows Aspose.Cells | header row AutoFitRows C# | workbook to workbook row transfer Aspose.Cells | maintain header height Aspose.Cells | C# Aspose.Cells copy rows and AutoFitRows

using System;
using System.IO;
using Aspose.Cells;

// The example creates a source workbook with header and data rows, copies the header rows and the remaining data rows separately into a destination workbook, applies AutoFitRows only to the header rows (rows 0‑1) to keep their height consistent, optionally auto‑fits the other rows, and saves the result as an XLSX file on the desktop.
class CopyRowsAndAutoFitHeader
{
    static void Main()
    {
        // Create source workbook and populate it with header and data rows
        Workbook sourceWorkbook = new Workbook();
        Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
        sourceSheet.Cells["A1"].PutValue("Header1");
        sourceSheet.Cells["B1"].PutValue("Header2");
        sourceSheet.Cells["A2"].PutValue("Header3");
        sourceSheet.Cells["B2"].PutValue("Header4");
        sourceSheet.Cells["A3"].PutValue("Data1");
        sourceSheet.Cells["B3"].PutValue("Data2");
        sourceSheet.Cells["A4"].PutValue("Data3");
        sourceSheet.Cells["B4"].PutValue("Data4");

        // Create destination workbook where rows will be copied
        Workbook destinationWorkbook = new Workbook();
        Worksheet destinationSheet = destinationWorkbook.Worksheets[0];

        // Copy header rows (rows 0 and 1) from source to destination
        // Parameters: source cells, source start row, destination start row, number of rows to copy
        destinationSheet.Cells.CopyRows(sourceSheet.Cells, 0, 0, 2);

        // Copy the remaining data rows starting from source row index 2
        int totalRows = sourceSheet.Cells.MaxDisplayRange.RowCount;
        int dataRowsCount = totalRows - 2; // rows after the header
        destinationSheet.Cells.CopyRows(sourceSheet.Cells, 2, 2, dataRowsCount);

        // AutoFit only the header rows (0 to 1) to keep their height consistent
        destinationSheet.AutoFitRows(0, 1);

        // Optionally, AutoFit the rest of the rows (2 to last row)
        destinationSheet.AutoFitRows(2, destinationSheet.Cells.MaxDisplayRange.RowCount - 1);

        // Save the resulting workbook
        string outputPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
            "CopiedWithHeaderAutoFit.xlsx");
        destinationWorkbook.Save(outputPath, SaveFormat.Xlsx);
    }
}
