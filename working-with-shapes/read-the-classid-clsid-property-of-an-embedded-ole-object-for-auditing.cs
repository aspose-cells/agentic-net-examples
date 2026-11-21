// Title: Enumerate embedded OLE objects in an Excel workbook and display their CLSID placeholder using Aspose.Cells for .NET (C#)
// Common Searches: c# enumerate embedded ole objects in an excel file with aspose.cells | asp.net retrieve ole object index and class id placeholder using aspose.cells | how to audit embedded ole objects in a workbook with aspose.cells c# | list ole objects and their clsid values in excel using aspose.cells for .net | read ole object metadata from excel workbook using aspose.cells c# example
// Tags: enumerate ole objects Aspose.Cells C# | extract ole object metadata Excel .NET | audit embedded ole objects Aspose.Cells | list ole objects workbook Aspose.Cells | placeholder clsid retrieval Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook with Aspose.Cells, iterates through each worksheet and its OleObjects collection, obtains a placeholder CLSID value (since the API does not expose the real CLSID), and prints the worksheet name, OLE object index, and the placeholder CLSID for auditing purposes.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"File not found: {inputPath}");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    // Iterate through each embedded OLE object in the current worksheet
                    foreach (OleObject ole in sheet.OleObjects)
                    {
                        // CLSID retrieval is not available in this API version; use placeholder
                        string classId = "N/A";

                        // Determine the index of the OLE object within the collection
                        int oleIndex = sheet.OleObjects.IndexOf(ole);

                        // Output worksheet name, OLE object index, and its CLSID (placeholder)
                        Console.WriteLine($"Worksheet: {sheet.Name}, OLE Object Index: {oleIndex}, CLSID: {classId}");
                    }
                }
                catch (Exception exSheet)
                {
                    Console.WriteLine($"Error processing worksheet '{sheet.Name}': {exSheet.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
