// Title: C# – Set Multiline Comments Built‑In Document Property in Aspose.Cells Workbook
// Description: Demonstrates how to create a new Aspose.Cells workbook, access its BuiltInDocumentPropertyCollection, assign a multiline string to the Comments property using Environment.NewLine (including author and date), display the value, and save the file as an .xlsx document.
// Keywords: Aspose.Cells | C# | set Comments property | built‑in document properties | multiline comments | Environment.NewLine | Excel metadata | Workbook comments | .NET Excel library | Aspose.Cells example
// Common Searches: Aspose.Cells set multiline Comments property C# | How to add line breaks to workbook Comments in Aspose.Cells | C# built‑in document properties Aspose.Cells example | Save custom notes in Excel file using Aspose.Cells | Aspose.Cells Comments metadata with date and author
// Developer Intent: Add a detailed, multiline comment to a workbook’s built‑in Comments property and persist it in the saved Excel file.
// Use Cases: Include audit‑trail notes (author, date, purpose) directly in the workbook metadata. | Provide end‑users with multi‑line documentation or usage instructions embedded in the file. | Store version‑specific remarks for automatically generated reports without altering worksheet content.
// AI Prompts: Generate C# code that reads the Comments built‑in property from an existing Aspose.Cells workbook and appends additional lines. | Show how to set other built‑in properties (Title, Subject, Keywords) together with a multiline Comments field in one snippet. | Explain how to preserve Unicode characters and line‑break formatting when exporting the Comments property to Excel with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

// Demonstrates how to create a new Aspose.Cells workbook, access its BuiltInDocumentPropertyCollection, assign a multiline string to the Comments property using Environment.NewLine (including author and date), display the value, and save the file as an .xlsx document.
class SetWorkbookComments
{
    static void Main()
    {
        // Create a new workbook (lifecycle: create)
        Workbook workbook = new Workbook();

        // Access the built‑in document properties
        BuiltInDocumentPropertyCollection properties = workbook.BuiltInDocumentProperties;

        // Set the Comments property with a multiline description
        properties.Comments = "This workbook was generated programmatically." + Environment.NewLine +
                             "It contains sample data for demonstration purposes." + Environment.NewLine +
                             "Author: John Doe" + Environment.NewLine +
                             "Date: " + DateTime.Now.ToString("yyyy-MM-dd");

        // Optionally display the comments to verify
        Console.WriteLine("Workbook Comments:");
        Console.WriteLine(properties.Comments);

        // Save the workbook (lifecycle: save)
        workbook.Save("WorkbookWithComments.xlsx");
    }
}
