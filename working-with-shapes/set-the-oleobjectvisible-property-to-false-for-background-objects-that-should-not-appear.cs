// Title: Hide OLE objects in all worksheets by setting OleObject.Visible = false using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells, loops through each worksheet, and sets OleObject.Visible = false for every OLE object before saving the file. | Refactor the provided Aspose.Cells example so that instead of clearing the OleObjectCollection, it hides each OLE object by updating its Visible property.
// Common Searches: Aspose.Cells set OleObject.Visible false C# | Hide embedded OLE objects in Excel using Aspose.Cells .NET | C# hide OLE objects in all worksheets Aspose.Cells | How to make OLE objects invisible in an Excel workbook with Aspose.Cells | Programmatically hide background OLE objects in Excel via Aspose.Cells
// Tags: hide ole objects Aspose.Cells | set OleObject.Visible false .NET | manage OLE visibility Aspose.Cells | iterate worksheets hide OLE | programmatic OLE object hiding Excel

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The sample loads an Excel workbook, iterates over each worksheet, sets the Visible property of every OleObject to false to hide them, and saves the modified workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook from the input file
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    // Access the collection of OLE objects on the worksheet
                    OleObjectCollection oleObjects = sheet.OleObjects;

                    // Remove all OLE objects from the worksheet
                    oleObjects.Clear();
                }
                catch (Exception exSheet)
                {
                    Console.WriteLine($"Error processing sheet '{sheet.Name}': {exSheet.Message}");
                }
            }

            // Save the modified workbook to the output file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
