// Title: Delete the first ActiveX (OLE) control from an Excel worksheet and confirm its removal with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells that opens a workbook, removes the first OleObject (ActiveX control) from a specified worksheet, checks that the OleObjects collection is empty, and saves the modified file. | Show how to output a console message indicating whether any ActiveX controls remain after calling Worksheet.OleObjects.RemoveAt in a .NET application. | Provide error‑handling snippets for missing input files and save failures when deleting OLE objects with Aspose.Cells.
// Common Searches: aspnet remove first ActiveX control from Excel worksheet using Aspose.Cells | how to verify OleObjects collection is empty after deletion in C# | Aspose.Cells delete OLE object and check remaining controls | C# console app remove ActiveX from .xlsx and confirm removal | sample code for removing ActiveX (OLE) objects with Aspose.Cells .NET
// Tags: remove ActiveX OleObject Aspose.Cells | verify OleObjects collection empty C# | delete OLE control from Excel worksheet | Aspose.Cells check for remaining ActiveX controls | C# workbook save after OLE removal

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook, removes the first ActiveX (OLE) object from the first worksheet if present, validates that no OleObjects remain, reports the result via console messages, and saves the updated workbook while handling missing files and save errors.
class RemoveActiveXControl
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook containing the ActiveX (OLE) control
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index or name as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Remove the first OLE (ActiveX) object if any exist
            if (sheet.OleObjects.Count > 0)
            {
                sheet.OleObjects.RemoveAt(0);
                Console.WriteLine("ActiveX control removed.");
            }
            else
            {
                Console.WriteLine("No ActiveX control found to remove.");
            }

            // Verify that no ActiveX controls remain in the worksheet
            bool anyActiveXRemaining = sheet.OleObjects.Count > 0;
            Console.WriteLine(anyActiveXRemaining
                ? "ActiveX control still present in the worksheet."
                : "ActiveX control successfully absent from the worksheet.");

            // Save the modified workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
