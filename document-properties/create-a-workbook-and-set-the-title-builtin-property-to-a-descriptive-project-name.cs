// Title: Create an Excel workbook and assign a custom Title built-in document property using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that creates a new workbook, sets the Title built-in document property to a specific project name, and saves the file. | Show how to open an existing Excel file with Aspose.Cells and update its Title property without affecting other data. | Provide a concise example that sets multiple built-in properties such as Title and Author in a workbook using Aspose.Cells.
// Common Searches: Aspose.Cells C# set workbook Title built-in property example | how to change Excel file title metadata programmatically with Aspose.Cells | C# code to add custom Title to Excel document using Aspose.Cells library | setting built-in document properties in a new workbook with Aspose.Cells for .NET | save Excel workbook after updating Title property using Aspose.Cells
// Tags: Aspose.Cells set workbook Title property | C# built-in document properties Excel | create workbook assign metadata Aspose | save Excel file after modifying built-in properties | Excel metadata Title Aspose.Cells

using Aspose.Cells;
using System;

// // Creates a new Excel workbook, assigns a descriptive Title built-in document property (e.g., "Project Alpha – Financial Forecast"), and saves the file as ProjectAlpha.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set the Title built‑in property to a descriptive project name
        workbook.BuiltInDocumentProperties.Title = "Project Alpha – Financial Forecast";

        // Save the workbook to a file (optional)
        workbook.Save("ProjectAlpha.xlsx");
    }
}
