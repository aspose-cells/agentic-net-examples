// Title: Set the Title built‑in document property of an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set workbook built‑in Title property programmatically | How to add a descriptive project name to Excel file metadata with Aspose.Cells | C# example for updating Excel built‑in document properties using Aspose.Cells | Saving an Excel workbook with a custom Title property in .NET | Aspose.Cells set built‑in document property Title before saving workbook
// Tags: Aspose.Cells set built-in Title property | C# update Excel document metadata | Excel workbook built-in properties manipulation | save workbook with custom Title metadata | Aspose.Cells document properties API

using Aspose.Cells;
using System;

// Creates a new Excel workbook, assigns "Descriptive Project Name" to the Title built‑in document property via Aspose.Cells, and saves the file as ProjectWorkbook.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set the Title built‑in property to a descriptive project name
        workbook.BuiltInDocumentProperties["Title"].Value = "Descriptive Project Name";

        // Save the workbook to a file
        workbook.Save("ProjectWorkbook.xlsx");
    }
}
