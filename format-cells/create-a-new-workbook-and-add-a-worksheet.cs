// Title: C# – Create Excel Workbook, Add Named Worksheet, Write Cell, and Save as XLSX with Aspose.Cells
// Description: Demonstrates how to instantiate an Aspose.Cells Workbook in C#, add a worksheet with a custom name, write a value to cell A1, and save the file as an XLSX document.
// Keywords: Aspose.Cells C# create workbook | Aspose.Cells add worksheet | Aspose.Cells save as xlsx | Aspose.Cells write cell value | Workbook Worksheets.Add | Aspose.Cells .NET example
// Common Searches: how to create a new Excel workbook with Aspose.Cells C# | add a worksheet with custom name using Aspose.Cells .NET | save Aspose.Cells workbook as xlsx file | write value to cell A1 with Aspose.Cells | Aspose.Cells example for creating template workbook
// Developer Intent: Generate a new Excel file, add a custom‑named sheet, insert initial data, and save it as XLSX using Aspose.Cells for .NET.
// Use Cases: Create a blank template with a predefined sheet name for downstream data import. | Automate report generation where each report starts with a header row on a specific worksheet. | Initialize workbooks in a web API before populating them with dynamic content. | Build Excel files for export from a desktop application with a default sheet layout.
// AI Prompts: Write C# code using Aspose.Cells to create a workbook, add a worksheet named "Data", place "Report" in cell B2, and save as "Report.xlsx". | Show how to add multiple worksheets with custom names, set column widths, and populate header rows before saving the file with Aspose.Cells .NET. | Explain step‑by‑step how to insert values into several cells on a newly added worksheet and then export the workbook as XLSX.

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Demonstrates how to instantiate an Aspose.Cells Workbook in C#, add a worksheet with a custom name, write a value to cell A1, and save the file as an XLSX document.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook instance (default format is Xlsx)
            Workbook workbook = new Workbook();

            // Add a new worksheet with a custom name to the workbook
            // The Add method returns the created Worksheet object
            Worksheet newSheet = workbook.Worksheets.Add("MyWorksheet");

            // Optionally, put some data into the new worksheet
            newSheet.Cells["A1"].PutValue("Hello, Aspose.Cells!");

            // Save the workbook to a file in the current directory
            // SaveFormat.Xlsx specifies the output file type
            workbook.Save("MyWorkbook.xlsx", SaveFormat.Xlsx);

            Console.WriteLine("Workbook created and saved successfully.");
        }
    }
}
