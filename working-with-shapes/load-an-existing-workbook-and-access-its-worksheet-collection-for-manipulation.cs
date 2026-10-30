// Title: C# – Load an existing Excel workbook and rename the first worksheet with Aspose.Cells
// Common Searches: how to open an existing Excel file and rename the first sheet using Aspose.Cells for .NET | Aspose.Cells C# load workbook and change worksheet name example | rename worksheet after loading workbook with Aspose.Cells .NET code | C# code to load input.xlsx and set first worksheet name using Aspose.Cells | save modified workbook to a new file after renaming a sheet with Aspose.Cells
// Tags: load workbook from file Aspose.Cells | access worksheets collection Aspose.Cells | rename first worksheet Aspose.Cells | save workbook to new file Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample loads 'input.xlsx' into an Aspose.Cells Workbook, retrieves its Worksheets collection, renames the first sheet to "FirstSheet" if present, and then saves the updated workbook as 'output.xlsx'.
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
                worksheets[0].Name = "FirstSheet";
            }

            // (Optional) Save the workbook after manipulation
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
        }
    }
}
