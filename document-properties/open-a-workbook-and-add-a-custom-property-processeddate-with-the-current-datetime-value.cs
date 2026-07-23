// Title: Aspose.Cells C# – Add a Custom Document Property "ProcessedDate" with the Current DateTime
// Description: Demonstrates how to create or load an Excel workbook with Aspose.Cells for .NET, add a custom document property named "ProcessedDate" that stores the current system date and time (DateTime.Now), and save the file. Ideal for embedding processing timestamps directly into workbook metadata.
// Keywords: Aspose.Cells | C# | custom document property | ProcessedDate | DateTime.Now | Excel metadata | add property | timestamp | save workbook | .NET
// Common Searches: Aspose.Cells add custom property to Excel | C# set processed date in workbook metadata | How to store timestamp in Excel file using Aspose | Add DateTime property to Excel workbook with Aspose.Cells | Save Excel file with custom document properties .NET
// Developer Intent: Insert a "ProcessedDate" custom property containing the current date and time into an Excel workbook using Aspose.Cells.
// Use Cases: Log the exact moment a file was generated for audit trails. | Provide downstream systems with a freshness indicator for data processing pipelines. | Support versioning by embedding a processing timestamp directly in the workbook.
// AI Prompts: Generate C# code with Aspose.Cells that adds or updates a "ProcessedDate" custom property using DateTime.UtcNow and saves the workbook. | Write a reusable method that checks for an existing "ProcessedDate" property in a workbook and replaces it with the current time. | Create a script that loads any Excel file, adds a "ProcessedDate" timestamp property, and returns the path of the saved file.

using System;
using Aspose.Cells;

// Demonstrates how to create or load an Excel workbook with Aspose.Cells for .NET, add a custom document property named "ProcessedDate" that stores the current system date and time (DateTime.Now), and save the file. Ideal for embedding processing timestamps directly into workbook metadata.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Add a custom document property named "ProcessedDate" with the current date and time
        workbook.CustomDocumentProperties.Add("ProcessedDate", DateTime.Now);

        // Save the workbook to a file
        workbook.Save("ProcessedWorkbook.xlsx");
    }
}
