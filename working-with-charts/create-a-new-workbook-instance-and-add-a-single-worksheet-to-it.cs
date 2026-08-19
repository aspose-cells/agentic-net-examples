// Title: Create a Workbook and Add a Worksheet in C# using Aspose.Cells
// Description: This C# snippet demonstrates how to instantiate an Aspose.Cells Workbook, append a new worksheet with Worksheets.Add(), retrieve the sheet via the returned index, write a value to cell A1, and optionally save the workbook as an XLSX file.
// Keywords: Aspose.Cells C# | .NET Excel workbook | Workbook.Add worksheet | Worksheets.Add example | Create worksheet Aspose.Cells | PutValue cell A1 | save workbook Aspose.Cells | Excel file generation C# | Aspose.Cells API usage | add sheet index
// Common Searches: how to add a worksheet with Aspose.Cells .NET | Aspose.Cells C# create workbook and new sheet | retrieve index of newly added worksheet Aspose.Cells | write value to cell A1 using Aspose.Cells | save Aspose.Cells workbook as xlsx
// Developer Intent: Append a single worksheet to a newly created workbook programmatically.
// Use Cases: Initialize a workbook before populating a report that requires a dedicated summary sheet. | Add a worksheet and set header text in A1 for data export templates. | Programmatically generate multiple sheets in a loop, using the returned index to customize each sheet's layout.
// AI Prompts: Generate C# code that creates an Aspose.Cells Workbook, adds three worksheets, and writes a unique title to cell A1 of each sheet. | Explain how to capture the index returned by Worksheets.Add() and use it for formatting, data insertion, and chart creation.

using System;
using Aspose.Cells;

// This C# snippet demonstrates how to instantiate an Aspose.Cells Workbook, append a new worksheet with Worksheets.Add(), retrieve the sheet via the returned index, write a value to cell A1, and optionally save the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Add a single worksheet to the workbook
        // The Add() method returns the index of the newly added sheet
        int newSheetIndex = workbook.Worksheets.Add();

        // Retrieve the worksheet object (optional, for further manipulation)
        Worksheet worksheet = workbook.Worksheets[newSheetIndex];

        // Example operation: put a value in cell A1 of the new worksheet
        worksheet.Cells["A1"].PutValue("Hello, Aspose.Cells!");

        // The workbook now contains the default worksheet plus the newly added one
        // (If you need to persist the file, uncomment the line below)
        // workbook.Save("CreatedWorkbook.xlsx", SaveFormat.Xlsx);
    }
}
