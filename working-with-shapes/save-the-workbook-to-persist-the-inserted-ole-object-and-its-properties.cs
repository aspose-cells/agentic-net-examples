// Title: How to save an Excel workbook and retain embedded OLE objects using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# save workbook without losing embedded OLE objects | preserve OLE shape properties when exporting Excel to XLSX with Aspose.Cells | load password‑protected Excel file and keep OLE objects using Aspose.Cells .NET | C# example to copy OLE objects from one workbook to another with Aspose.Cells | saving Excel file with embedded OLE objects using Aspose.Cells SaveFormat.Xlsx
// Tags: save workbook with OLE objects Aspose.Cells | preserve embedded OLE shapes .NET | load protected Excel workbook Aspose.Cells | export to Xlsx retaining OLE objects | C# Aspose.Cells OLE object handling

using Aspose.Cells;
using System;
using System.IO;

// The sample loads an existing Excel file (optionally using LoadOptions for password‑protected workbooks) and saves it as a new .xlsx file with Aspose.Cells, ensuring that all embedded OLE objects and their properties are preserved, with robust exception handling.
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
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // LoadOptions can be used to supply a password if the workbook is protected
            LoadOptions loadOptions = new LoadOptions();
            // loadOptions.Password = "yourPassword"; // Uncomment and set if needed

            // Load the workbook (with OLE objects) using the specified options
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Save the workbook, preserving OLE objects and their properties
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (CellsException ex)
        {
            // Handles Aspose.Cells specific errors (e.g., missing password)
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Handles any other unexpected errors
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
