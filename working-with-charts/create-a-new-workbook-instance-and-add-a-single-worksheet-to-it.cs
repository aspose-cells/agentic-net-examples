// Title: Aspose.Cells .NET: Create a Workbook and Add a Single Worksheet (C#)
// Description: Demonstrates how to instantiate a Workbook, add one Worksheet, optionally rename it, and save the file as an XLSX document using Aspose.Cells for .NET.
// Keywords: Aspose.Cells create workbook C# | add worksheet Aspose.Cells | Aspose.Cells set worksheet name | save Excel file Aspose.Cells | C# generate XLSX | Aspose.Cells API example | empty Excel workbook Aspose
// Common Searches: how to create a workbook with Aspose.Cells .NET | add a worksheet to an Aspose.Cells workbook | rename worksheet Aspose.Cells C# | save empty Excel file using Aspose.Cells | Aspose.Cells single sheet example
// Developer Intent: Generate a blank Excel file with one named sheet and persist it to disk.
// Use Cases: Create a template workbook that will later be populated with data. | Automate the production of a placeholder .xlsx file for downstream processes. | Build a simple one‑sheet report before adding charts or tables.
// AI Prompts: Write C# code using Aspose.Cells to create a workbook, add a worksheet called "Data", and save it as "Report.xlsx". | Show how to rename a newly added worksheet in Aspose.Cells and export the workbook to an XLSX file. | Provide a step‑by‑step explanation of creating an empty Excel file with a single sheet using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Demonstrates how to instantiate a Workbook, add one Worksheet, optionally rename it, and save the file as an XLSX document using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Add a single worksheet to the workbook
        int sheetIndex = workbook.Worksheets.Add();

        // Optionally set a name for the new worksheet
        workbook.Worksheets[sheetIndex].Name = "MySheet";

        // Save the workbook to a file
        workbook.Save("SingleWorksheet.xlsx");
    }
}
