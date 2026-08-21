// Title: Log OLE Object File Name and Size in an Excel Workbook with Aspose.Cells for .NET (C#)
// Description: This example shows how to open an Excel file, iterate through every worksheet, enumerate each OLE object, capture its original source file name (or a placeholder for embedded objects) and its byte size, output the details to the console for traceability, and finally save the workbook unchanged.
// Keywords: Aspose.Cells | C# | OLE object logging | Excel OLE source file name | OLE object size | ObjectSourceFullName | ObjectData length | traceability | .NET Excel automation | enumerate OLE objects
// Common Searches: Aspose.Cells log OLE object name and size | C# get OLE object source file Excel | how to read OLE object data length with Aspose.Cells | enumerate OLE objects in each worksheet .NET | trace embedded OLE objects in Excel file
// Developer Intent: Record the original file name and byte size of every OLE object in an Excel workbook to create an audit trail.
// Use Cases: Generate an audit log of all OLE objects before modifying a workbook. | Validate that embedded OLE objects stay within a size threshold. | Produce a report of linked OLE objects and their source files for documentation.
// AI Prompts: Write C# code using Aspose.Cells that extracts and logs the source file name and data size of each OLE object in an Excel workbook. | Create a method that returns a list of OLE object metadata (worksheet, index, source file, size) from a Workbook object. | Explain how to handle empty ObjectSourceFullName values for embedded OLE objects while logging their details.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This example shows how to open an Excel file, iterate through every worksheet, enumerate each OLE object, capture its original source file name (or a placeholder for embedded objects) and its byte size, output the details to the console for traceability, and finally save the workbook unchanged.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook that contains OLE objects
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                // Get the collection of OLE objects on the current worksheet
                OleObjectCollection oleObjects = worksheet.OleObjects;

                // Process each OLE object
                for (int i = 0; i < oleObjects.Count; i++)
                {
                    OleObject ole = oleObjects[i];

                    // Retrieve the original source file name.
                    // For linked OLE objects use ObjectSourceFullName.
                    // For embedded objects the property may be empty; use a placeholder.
                    string sourceFileName = ole.ObjectSourceFullName;
                    if (string.IsNullOrEmpty(sourceFileName))
                    {
                        sourceFileName = "[Embedded Object]";
                    }

                    // Determine the size of the OLE object's data (in bytes)
                    long dataSize = ole.ObjectData != null ? ole.ObjectData.Length : 0;

                    // Log the information for traceability
                    Console.WriteLine($"Worksheet: {worksheet.Name}, OLE Index: {i}, Source File: {sourceFileName}, Size: {dataSize} bytes");
                }
            }

            // Save the workbook (unchanged) to a new file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
