// Title: Rename the display label of linked OLE objects in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Loop through each worksheet with Aspose.Cells, access every OleObject, and assign a custom string to its Name property. | Programmatically change the display label of linked OLE objects in a workbook and save the updated file using C# and Aspose.Cells.
// Common Searches: C# Aspose.Cells change OleObject display label in Excel workbook | How to set Name property for linked OLE objects using Aspose.Cells | Iterate worksheets and rename OLE objects programmatically with Aspose.Cells .NET | Update OLE object label in existing Excel file Aspose.Cells example | Rename linked OLE object in Excel using Aspose.Cells C# code
// Tags: set OleObject.Name Aspose.Cells | rename linked OLE object Excel C# | iterate worksheets OleObject collection Aspose.Cells | update OLE display label Aspose.Cells | modify OLE object name in workbook

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an existing Excel file, iterates through all worksheets and their OleObject collections, assigns a custom descriptive string to each OLE object's Name property, and saves the workbook, handling missing files and runtime errors.
class UpdateOleObjectLabel
{
    static void Main()
    {
        const string inputFile = "Input.xlsx";
        const string outputFile = "Output.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Error: Input file \"{inputFile}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputFile);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Access the collection of OLE objects on the current worksheet
                OleObjectCollection oleObjects = sheet.OleObjects;

                // Loop through each OLE object
                foreach (OleObject ole in oleObjects)
                {
                    // Update the display label (Name property) for OLE objects
                    // If you need to differentiate between embedded and linked objects,
                    // use the appropriate property available in your Aspose.Cells version.
                    ole.Name = "Descriptive Label for Linked OLE";
                }
            }

            // Save the modified workbook
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully as \"{outputFile}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime exceptions
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
