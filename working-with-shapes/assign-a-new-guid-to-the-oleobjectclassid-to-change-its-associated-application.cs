// Title: Assign a new GUID to an OleObject's Name property in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an existing .xlsx file with Aspose.Cells, retrieves the first OleObject on the first worksheet, sets its Name to a newly generated GUID, and saves the workbook. | Show how to programmatically replace the identifier of an embedded OLE object in Excel by assigning a fresh GUID to the OleObject.Name using the Aspose.Cells .NET API.
// Common Searches: how to set a new GUID for an OleObject name with Aspose.Cells C# | Aspose.Cells change OleObject identifier in an Excel file | C# generate GUID and assign to embedded OLE object using Aspose.Cells | update OleObject Name property programmatically in .NET workbook
// Tags: Aspose.Cells set OleObject Name GUID | C# update OleObject identifier Aspose.Cells | modify embedded OLE object name .NET | generate GUID for Excel OLE object Aspose.Cells | change OleObject application reference via GUID

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program loads an existing Excel workbook, checks for at least one OleObject on the first worksheet, assigns a newly generated GUID string to the OleObject's Name property, and saves the modified file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one OleObject on the sheet
            if (sheet.OleObjects.Count > 0)
            {
                // Access the first OleObject
                OleObject ole = sheet.OleObjects[0];

                // Assign a new GUID (as string) to the object's name as an example modification
                ole.Name = Guid.NewGuid().ToString();
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
