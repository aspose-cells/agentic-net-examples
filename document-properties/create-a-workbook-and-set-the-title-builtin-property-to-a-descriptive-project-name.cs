// Title: How to set the Title built‑in document property in an Aspose.Cells workbook (C#)
// Description: This C# example shows how to create a new Workbook with Aspose.Cells, assign a descriptive project name to the built‑in Title property, and save the workbook as an XLSX file. Embedding the Title metadata improves document management, searchability, and compliance.
// Keywords: Aspose.Cells C# set Title property | Excel built‑in document properties | Workbook metadata Aspose | set workbook title Aspose.Cells | C# Excel file metadata | Aspose.Cells document properties example
// Common Searches: Aspose.Cells C# set workbook Title property | How to add Title metadata to Excel file using Aspose.Cells | C# code to set built‑in document properties in Excel | Set Excel file Title with Aspose.Cells .NET | Update Excel document Title programmatically
// Developer Intent: Assign a descriptive Title to a new workbook’s built‑in document property and save the file.
// Use Cases: Insert a project name as the Title before distributing quarterly reports. | Add Title metadata to improve search results in document management systems. | Automate generation of multiple workbooks with consistent Title information.
// AI Prompts: Generate C# code using Aspose.Cells to set several built‑in document properties (Title, Author, Subject) and save the workbook. | Provide an example that reads the Title property from an existing Excel file with Aspose.Cells. | Explain how to update the Title property of a workbook without modifying its worksheets.

using System;
using Aspose.Cells;

// This C# example shows how to create a new Workbook with Aspose.Cells, assign a descriptive project name to the built‑in Title property, and save the workbook as an XLSX file. Embedding the Title metadata improves document management, searchability, and compliance.
class Program
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Set the Title built‑in document property to a descriptive project name
        workbook.BuiltInDocumentProperties.Title = "Project XYZ – Quarterly Report";

        // Save the workbook to a file (XLSX format)
        workbook.Save("ProjectWorkbook.xlsx", SaveFormat.Xlsx);
    }
}
