// Title: Replace an embedded OLE object with a new Word document and modify its properties using Aspose.Cells for .NET
// AI Prompts: Load an existing Excel workbook, locate the first OleObject, assign the binary content of a .docx file to its ObjectData, set the Name and IsAutoSize flags, then save the file with Aspose.Cells. | Read a Word document into a byte array and update an OleObject's ObjectData, Name, and IsAutoSize properties in C# using the Aspose.Cells API.
// Common Searches: Aspose.Cells C# replace OLE object source file with a Word document | How to update OleObject Name and AutoSize properties in an Excel workbook using .NET | Change embedded Word file inside an Excel sheet programmatically with Aspose.Cells | Set ObjectData of an OleObject from a .docx file in C# | Replace first OLE object in worksheet and save workbook using Aspose.Cells for .NET
// Tags: replace OleObject data with docx Aspose.Cells | set OleObject Name and IsAutoSize C# | update embedded Word document in Excel worksheet | load workbook modify OLE object Aspose | assign byte array to OleObject ObjectData

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The sample loads 'input.xlsx', finds the first OleObject on the first worksheet, replaces its embedded content with the binary data of 'NewDocument.docx', updates the object's Name and IsAutoSize properties, and saves the modified workbook as 'output.xlsx', handling missing files and runtime errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string newDocPath = "NewDocument.docx";
            const string outputPath = "output.xlsx";

            // Verify that the required files exist before proceeding
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input workbook not found: {inputPath}");
            if (!File.Exists(newDocPath))
                throw new FileNotFoundException($"Source OLE document not found: {newDocPath}");

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one OLE object on the sheet
            if (sheet.OleObjects.Count > 0)
            {
                OleObject ole = sheet.OleObjects[0];

                // Replace the OLE object's data with the new Word document
                try
                {
                    byte[] docData = File.ReadAllBytes(newDocPath);
                    ole.ObjectData = docData;
                    // The file format is inferred from the data; explicit setting is not required
                }
                catch (Exception readEx)
                {
                    Console.WriteLine($"Failed to read new document: {readEx.Message}");
                    throw;
                }

                // Update additional properties as required
                ole.Name = "WordDocumentOle";
                ole.IsAutoSize = true;
                // ole.IconFile = "icon.ico"; // optionally set a custom icon file
            }
            else
            {
                Console.WriteLine("No OLE objects found on the worksheet.");
            }

            // Save the modified workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
                throw;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
