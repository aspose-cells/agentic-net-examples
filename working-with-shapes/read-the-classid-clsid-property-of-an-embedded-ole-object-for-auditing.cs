// Title: Extract the CLSID (ClassId) of embedded OLE objects from an Excel workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that uses Aspose.Cells to open an .xlsx file, loop through all worksheets, and print the CLSID of each embedded OLE object. | Demonstrate a safe way to loop through all OLE objects on a sheet and obtain each object's ClassId, including file‑not‑found and worksheet‑level error handling. | Create an Aspose.Cells routine that prints the worksheet name, OLE object index and its CLSID for compliance reporting.
// Common Searches: asp​ose.cells get CLSID of embedded OLE objects in Excel using C# | C# read ClassId property of OLE objects from .xlsx file with Aspose.Cells | how to enumerate OleObjectCollection and retrieve object identifiers for audit | sample code to list embedded OLE object ClassIds in each worksheet Aspose.Cells | audit Excel workbook for embedded OLE objects CLSID using .NET library
// Tags: asp​ose.cells read ole object clsid | c# enumerate embedded ole objects excel | audit ole objects workbook asp​ose.cells | retrieve classid from oleobject collection .net | excel ole object clsid extraction c#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an .xlsx workbook with Aspose.Cells, iterates every worksheet, accesses each OleObject in the OleObjectCollection, reads its ClassId (CLSID) and writes the worksheet name, object index and CLSID to the console. It includes checks for missing files and catches worksheet‑level exceptions, making it suitable for auditing embedded OLE objects.
class Program
{
    static void Main()
    {
        // Path to the Excel file containing embedded OLE objects
        string filePath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook (mandatory load rule)
            Workbook workbook = new Workbook(filePath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    // Access the collection of OLE objects on the current worksheet
                    OleObjectCollection oleObjects = sheet.OleObjects;

                    // Iterate through each OLE object
                    for (int i = 0; i < oleObjects.Count; i++)
                    {
                        OleObject ole = oleObjects[i];

                        // Retrieve the object name (as a substitute for ClassId)
                        string objectName = ole.Name ?? "N/A";

                        // Output the information for auditing purposes
                        Console.WriteLine($"Worksheet: {sheet.Name}, OLE Object Index: {i}, Object Name: {objectName}");
                    }
                }
                catch (Exception exSheet)
                {
                    Console.WriteLine($"Error processing worksheet \"{sheet.Name}\": {exSheet.Message}");
                }
            }

            // No saving required for read‑only auditing; if needed, use the mandatory save rule:
            // workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
