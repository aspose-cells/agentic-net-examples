// Title: C# Sample: Create a Workbook, Add Data to a Column, and Auto‑Fit Column Width with Aspose.Cells
// Description: This Aspose.Cells for .NET example demonstrates how to instantiate a Workbook, write three text values into column A of the first worksheet, automatically adjust the column width to fit the longest entry, and save the file as AutoFitColumnDemo.xlsx. The code is ready to run in any C# environment.
// Keywords: Aspose.Cells C# auto fit column | create workbook Aspose.Cells .NET | populate Excel column with data | auto‑fit column width example | save Excel file Aspose.Cells | C# Excel automation | Aspose.Cells sample code | GitHub Aspose.Cells demo
// Common Searches: Aspose.Cells auto fit column after adding data | C# example to create workbook and auto‑size columns | How to adjust Excel column width with Aspose.Cells .NET | Sample code for auto‑fitting columns in Aspose.Cells | Aspose.Cells column width automation
// Developer Intent: Generate an Excel workbook, insert text into a column, and automatically resize that column to accommodate the longest string.
// Use Cases: Producing reports where column widths must adapt to variable text lengths for clear presentation. | Exporting dynamic data to Excel while ensuring columns are sized correctly without manual formatting. | Building reusable templates that automatically prevent text truncation by auto‑fitting columns.
// AI Prompts: Provide C# code using Aspose.Cells to create a workbook, fill column B with values, and auto‑fit the column. | Show how to auto‑fit multiple columns after populating them with data in Aspose.Cells for .NET. | Explain how to programmatically adjust column widths for all used columns in an Aspose.Cells workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsAutoFitColumnDemo
{
    // This Aspose.Cells for .NET example demonstrates how to instantiate a Workbook, write three text values into column A of the first worksheet, automatically adjust the column width to fit the longest entry, and save the file as AutoFitColumnDemo.xlsx. The code is ready to run in any C# environment.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add sample data to column A (zero‑based index 0)
            worksheet.Cells["A1"].PutValue("Short");
            worksheet.Cells["A2"].PutValue("A much longer text that should expand the column width");
            worksheet.Cells["A3"].PutValue("Medium length");

            // Auto‑fit column A
            worksheet.AutoFitColumn(0);

            // Save the workbook
            workbook.Save("AutoFitColumnDemo.xlsx");
        }
    }
}
