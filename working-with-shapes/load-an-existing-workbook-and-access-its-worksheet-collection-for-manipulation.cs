// Title: Load an existing Excel workbook and rename its first worksheet with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells load workbook and change first sheet name | How to rename the first worksheet after opening an Excel file using Aspose.Cells | Aspose.Cells .NET example for accessing WorksheetCollection and updating sheet name | Programmatically rename a sheet in an existing .xlsx using Aspose.Cells C# | Load input.xlsx, rename sheet to RenamedSheet, and save with Aspose.Cells
// Tags: load workbook Aspose.Cells C# | access worksheets collection Aspose.Cells | update sheet name Aspose.Cells | save workbook Aspose.Cells | modify Excel sheet name .NET

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The code loads 'input.xlsx' into an Aspose.Cells Workbook, accesses its Worksheets collection, renames the first worksheet to 'RenamedSheet' if present, and saves the modified file as 'output.xlsx'.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the existing Excel file
            string inputPath = "input.xlsx";

            // Load the workbook from the file
            Workbook workbook = new Workbook(inputPath);

            // Access the collection of worksheets in the workbook
            WorksheetCollection worksheets = workbook.Worksheets;

            // Example manipulation: rename the first worksheet
            if (worksheets.Count > 0)
            {
                worksheets[0].Name = "RenamedSheet";
            }

            // (Optional) Save the modified workbook to a new file
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
        }
    }
}
