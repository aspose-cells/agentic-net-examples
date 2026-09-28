// Title: How to set a custom display label for an embedded OLE object in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells to embed a DOCX file as an OLE object in cell A1 and assign a custom label via the OleObject.Label property in C#. | Create a new workbook, add an OLE object from a byte array, set its display label, and save the workbook as XLSX with Aspose.Cells.
// Common Searches: C# Aspose.Cells set OleObject.Label after embedding a file | How to change the display name of an OLE object in an Excel file with Aspose | Assign custom label to embedded DOCX OLE object using Aspose.Cells .NET | Aspose.Cells example for adding OLE object and modifying its label property | Save Excel workbook with labeled OLE object using Aspose.Cells C#
// Tags: Aspose.Cells OleObject label property | C# embed file as OLE object | add OLE object to Excel worksheet | custom OLE object display name | save Excel with labeled OLE

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example creates a new workbook, embeds a DOCX file as an OLE object in cell A1, sets the OleObject.Label to a custom string, and saves the workbook as an XLSX file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Path to the file that will be embedded as an OLE object
            string oleSourcePath = "sample.docx";

            // Ensure the source file exists before attempting to embed it
            if (!File.Exists(oleSourcePath))
            {
                Console.WriteLine($"Error: The file \"{oleSourcePath}\" was not found.");
            }
            else
            {
                try
                {
                    // Read the file into a byte array (required by the Add method)
                    byte[] oleData = File.ReadAllBytes(oleSourcePath);

                    // Add an OLE object to cell A1 (row 0, column 0) with a size of 100x100 pixels
                    // The Add method returns the index of the created OleObject
                    int oleIndex = sheet.OleObjects.Add(0, 0, 100, 100, oleData);

                    // Retrieve the OleObject using the returned index
                    OleObject ole = sheet.OleObjects[oleIndex];

                    // Assign a custom display label to the OLE object
                    ole.Label = "My Custom Document";
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to embed OLE object: {ex.Message}");
                }
            }

            // Save the workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully as {outputPath}.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
