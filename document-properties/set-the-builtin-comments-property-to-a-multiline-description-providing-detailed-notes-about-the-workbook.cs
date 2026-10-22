// Title: Set a multiline Comments built‑in document property in an Excel workbook with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set Comments built‑in property with line breaks | How to add multiline notes to Excel workbook properties using Aspose.Cells | Saving workbook with custom Comments property containing date in Aspose.Cells .NET | Set built‑in document property 'Comments' to a multi‑line string in C# | Aspose.Cells example for writing detailed notes to Excel file metadata
// Tags: Aspose.Cells set built‑in Comments property | multiline Excel document property C# | write workbook metadata with line breaks | Aspose.Cells built‑in document properties example | Excel file comments property Aspose.Cells .NET

using Aspose.Cells;
using System;

// The sample creates a new Workbook, builds a multiline string that includes static text, the current date, and a review note, assigns this string to the workbook's built‑in Comments document property, and saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Define a multiline description for the Comments built‑in property
        string comments = "This workbook contains financial data.\r\n" +
                          "Generated on: " + DateTime.Now.ToString("yyyy-MM-dd") + "\r\n" +
                          "Please review the sheets before distribution.";

        // Set the built‑in Comments property to the multiline description
        workbook.BuiltInDocumentProperties["Comments"].Value = comments;

        // Save the workbook to a file
        workbook.Save("Output.xlsx");
    }
}
