// Title: AutoFit Header Row After Copying Rows with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to copy a header and data rows from one worksheet to another using Aspose.Cells, then apply AutoFitRows only to the header row to keep its height consistent while leaving other rows unchanged. The workbook is saved as an Excel file.
// Keywords: Aspose.Cells AutoFitRows C# | copy rows Excel Aspose.Cells | fit header row after copy | AutoFitRows specific range | preserve header height .NET | Excel row copy Aspose | C# Aspose.Cells example
// Common Searches: Aspose.Cells autofit only header row | C# copy rows then autofit header | how to keep header height after copying rows Aspose | AutoFitRows range parameters example | copy worksheet rows and adjust header height
// Developer Intent: Copy rows between worksheets and automatically adjust the height of the header row without affecting the rest of the data.
// Use Cases: Template‑based report generation where the header must retain a uniform height after copying data. | Creating a summary sheet that reuses a styled header from a source workbook while applying custom row heights to data rows later. | Building a reusable utility that copies a variable number of rows and applies AutoFitRows exclusively to the first row for consistent layout.
// AI Prompts: Write a C# function that copies rows from one Aspose.Cells worksheet to another and calls AutoFitRows only on the first row. | Show how to use AutoFitRows(startRow, endRow) to adjust just the header after copying a dynamic range of rows with Aspose.Cells. | Explain the steps to preserve header row height when copying tables between workbooks using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Demonstrates how to copy a header and data rows from one worksheet to another using Aspose.Cells, then apply AutoFitRows only to the header row to keep its height consistent while leaving other rows unchanged. The workbook is saved as an Excel file.
class AutoFitHeaderAfterCopy
{
    static void Main()
    {
        // ---------- Create source workbook and fill header + data ----------
        Workbook sourceWorkbook = new Workbook();
        Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

        // Header row (row 0)
        sourceSheet.Cells["A1"].PutValue("Header1");
        sourceSheet.Cells["B1"].PutValue("Header2");

        // Sample data rows
        sourceSheet.Cells["A2"].PutValue("Data1");
        sourceSheet.Cells["B2"].PutValue("Data2");
        sourceSheet.Cells["A3"].PutValue("Data3");
        sourceSheet.Cells["B3"].PutValue("Data4");

        // ---------- Create destination workbook ----------
        Workbook destinationWorkbook = new Workbook();
        Worksheet destinationSheet = destinationWorkbook.Worksheets[0];

        // ---------- Copy rows from source to destination ----------
        // Copy all rows that contain data (including header)
        int totalRowsToCopy = sourceSheet.Cells.MaxDisplayRange.RowCount; // e.g., 3 rows
        destinationSheet.Cells.CopyRows(sourceSheet.Cells, 0, 0, totalRowsToCopy);

        // ---------- AutoFit only the header rows ----------
        // Assuming header occupies the first row (index 0)
        destinationSheet.AutoFitRows(0, 0);

        // (Optional) AutoFit the remaining data rows separately
        // destinationSheet.AutoFitRows(1, totalRowsToCopy - 1);

        // ---------- Save the resulting workbook ----------
        destinationWorkbook.Save("CopiedWithHeaderAutoFit.xlsx");
    }
}
