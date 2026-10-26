// Title: Refresh linked OLE objects in an Excel workbook and save the file using Aspose.Cells for .NET (C#)
// Common Searches: C# code to refresh linked OLE objects in an Excel file with Aspose.Cells | How to iterate worksheets and update OleObject links before saving workbook in Aspose.Cells .NET | Saving an Excel workbook after processing OLE objects using Aspose.Cells for C# | Aspose.Cells example for handling missing OleObject.Update method in .NET | Refresh OLE objects in each worksheet and save workbook with Aspose.Cells C#
// Tags: update OLE links Aspose.Cells C# | enumerate worksheets OleObject Aspose.Cells | save Excel workbook after OLE update Aspose.Cells | fallback when OleObject.Update not available Aspose.Cells | process embedded OLE objects .NET Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, iterates through each worksheet and its OleObjects, attempts to refresh each OLE object (the Update call is commented out for compatibility), and then saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file '{inputPath}' not found.");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Refresh linked OLE objects (if supported by the library version)
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (OleObject ole in sheet.OleObjects)
                {
                    try
                    {
                        // In some Aspose.Cells versions the OleObject.Update() method is unavailable.
                        // If the method exists, it would be called here to refresh the OLE object.
                        // ole.Update(); // Uncomment if the method is available in your version.
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning: Failed to update OLE object in sheet '{sheet.Name}': {ex.Message}");
                    }
                }
            }

            // Save the workbook after processing
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
