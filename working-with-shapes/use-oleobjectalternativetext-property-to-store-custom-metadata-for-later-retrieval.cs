// Title: How to embed an OLE object in an Excel worksheet and store custom key‑value metadata using OleObject.AlternativeText with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# embed OLE object from byte array and set alternative text | store custom key value pairs in OleObject.AlternativeText in Excel using .NET | retrieve metadata from an OLE object in a saved workbook with Aspose.Cells | add OLE object to worksheet and read AlternativeText property later | save and load OLE object metadata in Excel file using Aspose.Cells C#
// Tags: add ole object from byte array Aspose.Cells | set oleobject alternativetext metadata | read oleobject alternativetext property | embed text file as ole object Excel .NET | save workbook with ole object Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates creating an OLE object from a text file, assigning custom key‑value metadata to its AlternativeText property, saving the workbook, then loading it back and retrieving the stored metadata using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // ---------- Create ----------
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the OLE source file exists; create a simple text file if missing
            string oleFilePath = "sample.txt";
            if (!File.Exists(oleFilePath))
            {
                File.WriteAllText(oleFilePath, "This is a sample OLE text file.");
            }

            // Read the OLE file bytes
            byte[] oleData = File.ReadAllBytes(oleFilePath);

            // Add an OLE object to the worksheet using the overload that accepts raw data
            int row = 1;
            int column = 1;
            int height = 100;
            int width = 200;

            // The overload requires: row, column, height, width, unknownInt, dataLength, dataBytes
            // The unknownInt parameter is not used for simple file embedding; pass 0.
            OleObject oleObject = worksheet.Shapes.AddOleObject(
                row, column, height, width, 0, oleData.Length, oleData);

            // Set display options (IconIndex and IconLabel are not available in this API version)
            oleObject.DisplayAsIcon = false;

            // Store custom metadata in the AlternativeText property
            oleObject.AlternativeText = "Author=John Doe;Version=1.0;Notes=Sample metadata";

            // Save the workbook
            string savePath = "OleMetadata.xlsx";
            workbook.Save(savePath, SaveFormat.Xlsx);

            // ---------- Load ----------
            // Verify the saved workbook exists before loading
            if (!File.Exists(savePath))
                throw new FileNotFoundException($"Saved workbook not found: {savePath}");

            Workbook loadedWorkbook = new Workbook(savePath);
            Worksheet loadedWorksheet = loadedWorkbook.Worksheets[0];

            // Find the first OLE object in the worksheet
            OleObject loadedOleObject = null;
            foreach (Shape shape in loadedWorksheet.Shapes)
            {
                if (shape is OleObject obj)
                {
                    loadedOleObject = obj;
                    break;
                }
            }

            if (loadedOleObject != null)
            {
                // Retrieve the stored metadata from AlternativeText
                string retrievedMetadata = loadedOleObject.AlternativeText;
                Console.WriteLine("Retrieved metadata: " + retrievedMetadata);
            }
            else
            {
                Console.WriteLine("No OLE object found in the loaded worksheet.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
