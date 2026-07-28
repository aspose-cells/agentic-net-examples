// Title: Create a New Workbook and Add a Named Worksheet with Aspose.Cells for .NET (C#)
// Description: This example shows how to instantiate a Workbook in the default XLSX format, add a worksheet called "MySheet", write "Hello Aspose.Cells" to cell A1, and save the file as CreatedWorkbook.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# create workbook | add worksheet Aspose.Cells | write cell value Aspose.Cells | save workbook as Xlsx | Excel automation C# | Aspose.Cells example
// Common Searches: Aspose.Cells create workbook C# | how to add worksheet with Aspose.Cells .NET | write value to cell A1 using Aspose.Cells | save Excel file as Xlsx Aspose.Cells
// Developer Intent: Generate an empty Excel file, insert a custom‑named sheet, place a value in A1, and persist the workbook as an XLSX document.
// Use Cases: Build a template workbook with a predefined sheet for automated reporting pipelines. | Create a fresh spreadsheet, set header cells, and deliver it to downstream services. | Provide a blank workbook that a web UI can later populate with user data.
// AI Prompts: Generate C# code with Aspose.Cells that creates a workbook, adds a sheet named "Data", writes the current date to B2, and saves as .xlsb. | Explain how to add multiple worksheets in a loop and set a default column width for each using Aspose.Cells for .NET. | Show an example of creating a workbook, inserting a chart sheet, and exporting the file to PDF with Aspose.Cells.

using System;
using Aspose.Cells;

// This example shows how to instantiate a Workbook in the default XLSX format, add a worksheet called "MySheet", write "Hello Aspose.Cells" to cell A1, and save the file as CreatedWorkbook.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Initialize a new workbook (default Xlsx format)
        Workbook workbook = new Workbook();

        // Add a new worksheet named "MySheet" to the workbook
        Worksheet worksheet = workbook.Worksheets.Add("MySheet");

        // Example: write a value into cell A1 of the new sheet
        worksheet.Cells["A1"].PutValue("Hello Aspose.Cells");

        // Save the workbook to a file
        workbook.Save("CreatedWorkbook.xlsx", SaveFormat.Xlsx);
    }
}
