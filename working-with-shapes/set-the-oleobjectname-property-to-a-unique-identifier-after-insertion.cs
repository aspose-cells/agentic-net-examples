// Title: Set a unique GUID as the Name of an OLE object after inserting it with Aspose.Cells for .NET
// Common Searches: how to assign a GUID to an OleObject name in Aspose.Cells C# | Aspose.Cells add OLE object and set a unique name property | C# generate a random identifier for an inserted OLE object using Aspose.Cells | set OleObject.Name to a new GUID after adding OLE object in .NET | example of naming OLE objects uniquely with Aspose.Cells
// Tags: OleObject.Name GUID assignment Aspose.Cells | add OLE object with unique identifier .NET | C# generate unique OleObject name | Aspose.Cells OLE object naming strategy | set unique name for inserted OLE object

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The code reads a DOCX file, creates a workbook, inserts the file as an OLE object at cell B2, assigns a newly generated GUID string to the OleObject's Name property, and saves the workbook as Result.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the OLE source file
            string oleFilePath = @"C:\Temp\sample.docx";

            // Verify that the OLE source file exists to avoid FileNotFoundException
            if (!File.Exists(oleFilePath))
            {
                Console.WriteLine($"The file '{oleFilePath}' was not found.");
                return;
            }

            // Read the OLE file into a byte array (required by AddOleObject overload)
            byte[] oleData;
            try
            {
                oleData = File.ReadAllBytes(oleFilePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to read OLE file: {ex.Message}");
                return;
            }

            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Insert an OLE object at cell B2 (row index 1, column index 1)
            // Height and width are in pixels; lower‑right cell coordinates are set to the same cell
            OleObject ole = sheet.Shapes.AddOleObject(
                1,               // upper‑left row index (zero‑based)
                1,               // upper‑left column index (zero‑based)
                100,             // height in pixels
                100,             // width in pixels
                1,               // lower‑right row index
                1,               // lower‑right column index
                oleData          // OLE data as byte array
            );

            // Set a unique identifier for the OLE object
            ole.Name = Guid.NewGuid().ToString();

            // Save the workbook (save rule)
            string resultPath = "Result.xlsx";
            workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved successfully as {resultPath}.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
