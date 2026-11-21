// Title: How to enumerate worksheets and obtain each sheet's internal index (SheetId) with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells loop through worksheets to read sheet index | Get worksheet IDs from an Excel file using Aspose.Cells .NET | Display sheet index and name while iterating workbook worksheets in C# | Aspose.Cells retrieve internal SheetId for diagnostic logging | Iterate workbook worksheets and output their indexes with Aspose.Cells
// Tags: Aspose.Cells worksheet index enumeration | C# iterate workbook worksheets | retrieve sheet Id Aspose.Cells | diagnostic logging Excel sheet indexes | load workbook and enumerate sheets .NET

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook with Aspose.Cells, iterates over the Worksheets collection, reads each worksheet's Index property (the internal SheetId), prints the index and sheet name for diagnostic purposes, and saves the workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Obtain the internal index of the current worksheet
                int sheetIndex = sheet.Index;

                // Diagnostic output: display Index and worksheet name
                Console.WriteLine($"Sheet Index: {sheetIndex}, Name: {sheet.Name}");
            }

            // Save the workbook (optional if modifications were made)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
