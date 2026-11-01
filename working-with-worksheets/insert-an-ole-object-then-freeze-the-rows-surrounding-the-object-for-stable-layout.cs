// Title: How to embed a DOCX file as an OLE object in an Excel worksheet and freeze the rows above it using Aspose.Cells for .NET (C#)
// AI Prompts: Read a DOCX file into a byte array, add it as an OLE object at row 5 column 2 with Aspose.Cells, then call FreezePanes to lock rows 0‑4 in the same worksheet. | Create a new workbook, insert a Word document as an OLE object with specific dimensions, and apply FreezePanes so the top five rows remain static while scrolling.
// Common Searches: C# Aspose.Cells embed Word document as OLE object and freeze top rows | Aspose.Cells FreezePanes after inserting OLE object example | How to add a DOCX OLE object to Excel using Aspose.Cells .NET | Freeze rows above an embedded OLE object in an Excel file with C# | Aspose.Cells add OLE object from byte array and set pane freeze
// Tags: insert Word document OLE object Aspose.Cells | freeze rows after OLE embed Aspose.Cells | add OLE object from byte array C# | worksheet FreezePanes API Aspose.Cells | read DOCX file to byte array C#

using System;
using System.IO;
using Aspose.Cells;

// The example verifies a DOCX file, reads it into a byte array, creates a new workbook, inserts the file as an OLE object at row 5 column 2 with defined height and width, freezes the first five rows using FreezePanes, and saves the workbook as OutputWithOleAndFrozenRows.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Verify that the OLE source file exists to avoid FileNotFoundException
            const string oleFilePath = "sample.docx";
            if (!File.Exists(oleFilePath))
            {
                Console.WriteLine($"Error: Required file \"{oleFilePath}\" not found.");
                return;
            }

            // Read the OLE file into a byte array (required by OleObjects.Add)
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

            // Create a new workbook
            var workbook = new Workbook();

            // Access the first worksheet
            var worksheet = workbook.Worksheets[0];

            // Add an OLE object (e.g., a Word document) at row 5, column 2.
            // The Add method returns the index of the newly added object.
            int oleIndex = worksheet.OleObjects.Add(
                5,          // upperLeftRow (0‑based)
                2,          // upperLeftColumn (0‑based)
                150,        // height in points
                200,        // width in points
                oleData     // file data as byte[]
            );

            // Retrieve the OleObject to adjust its properties if needed
            var oleObject = worksheet.OleObjects[oleIndex];
            // Example: you can modify display properties here
            // oleObject.Width = 200;
            // oleObject.Height = 150;

            // Freeze rows above the OLE object (rows 0‑4) while keeping columns unfrozen
            // Use the overload with four parameters: row, column, totalRows, totalColumns
            worksheet.FreezePanes(5, 0, 5, 0);

            // Save the workbook to a file
            const string outputPath = "OutputWithOleAndFrozenRows.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
