// Title: C# – Retrieve and Store a Worksheet by Name Using Aspose.Cells
// Description: Shows how to create a workbook, rename the default sheet, add a second sheet, then access a worksheet by its name with workbook.Worksheets["SheetName"] and keep the reference for further cell updates before saving the file. Example uses Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | retrieve worksheet by name | store worksheet reference | access worksheet Aspose | worksheet cell operations | save workbook Aspose.Cells | example code
// Common Searches: Aspose.Cells get worksheet by sheet name C# | How to access a worksheet using its name in Aspose.Cells .NET | Retrieve worksheet reference Aspose.Cells C# example | Write to a named worksheet with Aspose.Cells | Aspose.Cells workbook.Worksheets["SheetName"] usage
// Developer Intent: Obtain a specific worksheet by its name and retain the object for subsequent manipulation in C# with Aspose.Cells.
// Use Cases: Update header rows in a known sheet without looping through all worksheets. | Insert calculated values into cells of a sheet identified by name. | Copy data from a named worksheet to another sheet within the same workbook.
// AI Prompts: Generate C# code that opens an existing workbook, accesses the worksheet named "Report", writes a SUM formula to cell C5, and saves the file. | Provide an Aspose.Cells example that retrieves a sheet called "Log", writes the current timestamp to cell A2, and exports the workbook to PDF. | Show how to locate a worksheet by name, add a new column with incremental numbers, rename the sheet to "Processed", and save the workbook.

using System;
using Aspose.Cells;

namespace WorksheetAccessExample
{
    // Shows how to create a workbook, rename the default sheet, add a second sheet, then access a worksheet by its name with workbook.Worksheets["SheetName"] and keep the reference for further cell updates before saving the file. Example uses Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (default contains one worksheet)
            Workbook workbook = new Workbook();

            // Rename the default worksheet to a known name
            Worksheet firstSheet = workbook.Worksheets[0];
            firstSheet.Name = "DataSheet";

            // Add another worksheet for demonstration
            Worksheet secondSheet = workbook.Worksheets.Add("Summary");

            // Access the worksheet by its name and store the reference
            Worksheet accessedSheet = workbook.Worksheets["DataSheet"];

            // Perform further operations on the accessed worksheet
            accessedSheet.Cells["A1"].PutValue("Accessed by name successfully!");
            accessedSheet.Cells["A2"].PutValue(DateTime.Now);

            // Example: write something to the second sheet as well
            secondSheet.Cells["B1"].PutValue("This is the Summary sheet.");

            // Save the workbook to a file
            workbook.Save("WorksheetAccessDemo.xlsx");
        }
    }
}
