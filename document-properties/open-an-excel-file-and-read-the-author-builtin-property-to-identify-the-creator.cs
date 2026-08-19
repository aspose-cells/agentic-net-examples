// Title: Read Excel Workbook Author Property with Aspose.Cells for .NET (C#)
// Description: Loads an .xlsx file using Aspose.Cells, reads the BuiltInDocumentProperties.Author field, and prints the author name to the console.
// Keywords: Aspose.Cells | C# | read author property | built‑in document properties | Excel metadata | Workbook.Author | retrieve creator name | document properties .NET
// Common Searches: Aspose.Cells read author C# | Get Excel file author using Aspose | How to access built‑in document properties in .NET | C# retrieve workbook creator name | Aspose.Cells built‑in properties example
// Developer Intent: Extract the Author built‑in document property from an Excel workbook.
// Use Cases: Show the spreadsheet creator in a console tool for quick verification. | Log the author of uploaded Excel files to support audit trails in backend services. | Validate that a workbook originates from a specific user before initiating data processing. | Display author information in a UI dashboard that aggregates document metadata.
// AI Prompts: Generate C# code with Aspose.Cells that reads all built‑in document properties from an Excel file. | Show how to change the Author property of a workbook and save the file using Aspose.Cells. | Explain strategies for handling missing or empty Author values when reading Excel metadata with Aspose.Cells.

using System;
using Aspose.Cells;

// Loads an .xlsx file using Aspose.Cells, reads the BuiltInDocumentProperties.Author field, and prints the author name to the console.
class ReadAuthorProperty
{
    static void Main()
    {
        // Path to the Excel file to be opened
        string filePath = "input.xlsx";

        // Load the workbook from the specified file
        Workbook workbook = new Workbook(filePath);

        // Access the built‑in Author property of the workbook
        string author = workbook.BuiltInDocumentProperties.Author;

        // Display the author name
        Console.WriteLine("Author: " + author);
    }
}
