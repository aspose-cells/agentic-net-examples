// Title: Set a multiline Comments built‑in document property in an Excel workbook with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set Comments built‑in document property with line breaks | how to add multiline notes to Excel workbook using Aspose.Cells | using Environment.NewLine to create multiline document properties in Aspose.Cells | save Excel file with custom Comments property in .NET | populate built‑in Comments field of a workbook programmatically with Aspose.Cells
// Tags: set multiline built‑in document property Aspose.Cells | Aspose.Cells Comments property C# | environment.newline Excel document properties | save workbook with custom comments Aspose.Cells | populate built‑in Comments field .NET

using System;
using Aspose.Cells;

// Creates a new Workbook, builds a multiline string with Environment.NewLine, assigns it to the built‑in Comments property via workbook.BuiltInDocumentProperties["Comments"].Value, and saves the workbook as WorkbookWithComments.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Prepare a multiline description for the built‑in Comments property
        string comments = "Project: Quarterly Report" + Environment.NewLine +
                          "Author: Jane Doe" + Environment.NewLine +
                          "Reviewed by: John Smith" + Environment.NewLine +
                          "Notes: All figures are provisional and subject to change.";

        // Set the built‑in Comments property
        workbook.BuiltInDocumentProperties["Comments"].Value = comments;

        // Save the workbook to a file
        workbook.Save("WorkbookWithComments.xlsx");
    }
}
