// Title: Add a "ProcessedDate" custom document property (current DateTime) to an Excel workbook with Aspose.Cells for .NET
// Description: Creates or loads a Workbook, adds a custom document property named ProcessedDate containing DateTime.Now, and saves the workbook as an XLSX file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells custom document property | C# add Excel custom property | ProcessedDate property Aspose | Workbook.CustomDocumentProperties | save workbook with metadata | Excel metadata .NET
// Common Searches: how to add custom property to Excel using Aspose.Cells C# | Aspose.Cells set ProcessedDate property | add current datetime to workbook properties .NET | read custom document property Aspose.Cells | save Excel file with custom metadata using Aspose
// Developer Intent: Insert a custom document property called "ProcessedDate" that holds the current date and time into a workbook and persist the file.
// Use Cases: Timestamp generated reports for audit trails and traceability. | Embed processing dates to support version control and data lineage in exported spreadsheets. | Provide downstream automation with a reliable processing date stored in workbook metadata.
// AI Prompts: Generate C# code with Aspose.Cells that adds a "ProcessedDate" custom property using DateTime.UtcNow and saves the workbook as XLSX. | Show how to retrieve the "ProcessedDate" custom document property from an existing Excel file using Aspose.Cells. | Create robust error handling for adding or reading custom document properties when loading a workbook from a stream in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

// Creates or loads a Workbook, adds a custom document property named ProcessedDate containing DateTime.Now, and saves the workbook as an XLSX file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Add a custom document property named "ProcessedDate" with the current date and time
        workbook.CustomDocumentProperties.Add("ProcessedDate", DateTime.Now);

        // Save the workbook to a file
        workbook.Save("ProcessedWorkbook.xlsx", SaveFormat.Xlsx);
    }
}
