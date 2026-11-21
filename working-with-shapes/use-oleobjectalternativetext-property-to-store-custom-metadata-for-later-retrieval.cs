// Title: Embedding a Word document as an OLE object and storing custom JSON metadata with OleObject.AlternativeText using Aspose.Cells for .NET
// Common Searches: embed a Word file as an OLE object in an Excel worksheet using Aspose.Cells for .NET | save custom JSON string in OleObject.AlternativeText property with C# | read AlternativeText from an embedded OLE object in a workbook using Aspose.Cells | Aspose.Cells example for adding OLE object and attaching metadata | how to retrieve metadata from an OLE object in Excel via Aspose.Cells API
// Tags: Word document OLE embedding Aspose.Cells | OleObject AlternativeText JSON storage | C# retrieve OLE object metadata Excel | Aspose.Cells custom metadata in embedded objects | Excel OLE object handling with Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a new workbook, embeds a Word document as an OLE object, assigns a JSON string to the object's AlternativeText property, saves the file, then reloads the workbook, reads the AlternativeText from the OLE object, and outputs the stored metadata.
class OleObjectMetadataExample
{
    static void Main()
    {
        // Paths for the workbook and the file to embed
        string workbookPath = "OleObjectMetadata.xlsx";
        string embeddedFilePath = "sample.docx";

        try
        {
            // Ensure the embedded file exists
            if (!File.Exists(embeddedFilePath))
                throw new FileNotFoundException("Embedded file not found.", embeddedFilePath);

            // -------------------- Create and save workbook --------------------
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Read the file bytes to embed as an OLE object
            byte[] oleData = File.ReadAllBytes(embeddedFilePath);

            // Add the OLE object (positioned at row 5, column 2, spanning to row 10, column 5)
            // The Add method returns the index of the newly added object
            int oleIndex = sheet.OleObjects.Add(5, 2, 10, 5, oleData, "Word.Document.12");
            OleObject oleObject = sheet.OleObjects[oleIndex];

            // Store custom metadata in the AlternativeText property (e.g., JSON string)
            oleObject.AlternativeText = "{\"Author\":\"John Doe\",\"Version\":1,\"Notes\":\"Sample metadata\"}";

            // Save the workbook
            workbook.Save(workbookPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error during workbook creation: " + ex.Message);
            return;
        }

        try
        {
            // -------------------- Load and retrieve metadata --------------------
            // Ensure the workbook file exists before loading
            if (!File.Exists(workbookPath))
                throw new FileNotFoundException("Workbook file not found.", workbookPath);

            // Load the workbook from file
            Workbook loadedWorkbook = new Workbook(workbookPath);
            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];

            // Retrieve the first OLE object (assuming only one was added)
            if (loadedSheet.OleObjects.Count == 0)
                throw new InvalidOperationException("No OLE objects found in the worksheet.");

            OleObject loadedOleObject = loadedSheet.OleObjects[0];

            // Read the stored metadata from AlternativeText
            string storedMetadata = loadedOleObject.AlternativeText;

            // Output the retrieved metadata
            Console.WriteLine("Retrieved metadata: " + storedMetadata);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error during metadata retrieval: " + ex.Message);
        }
    }
}
