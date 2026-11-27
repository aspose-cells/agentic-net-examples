// Title: Save an Excel workbook after enumerating linked OLE objects with Aspose.Cells for .NET
// Common Searches: how to save a workbook after iterating OleObjectCollection using Aspose.Cells C# | Aspose.Cells enumerate linked OLE objects in an Excel file | C# code to check for linked OLE objects with Aspose.Cells | saving a new XLSX file after processing OLE objects in Aspose.Cells | why refresh linked OLE objects is not supported in Aspose.Cells .NET
// Tags: Aspose.Cells enumerate OleObjectCollection | save workbook Aspose.Cells Xlsx | linked OLE objects Excel Aspose.Cells | C# Aspose.Cells OLE object handling | Aspose.Cells workbook save after processing

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace RefreshLinkedOleObjects
{
    // The example loads an existing XLSX workbook using Aspose.Cells for .NET, iterates each worksheet's OleObjectCollection to log linked OLE objects (noting that direct refresh of linked OLE objects is currently unavailable), and then saves the workbook to a new file in Xlsx format with proper error handling.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths to the input and output files.
            // Adjust these paths to point to actual files on your system.
            string inputFile = "LinkedOleObjects.xlsx";
            string outputFile = "LinkedOleObjects_Refreshed.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException.
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file \"{inputFile}\" not found.");
                    return;
                }

                // Load the workbook that contains linked OLE objects.
                Workbook workbook = new Workbook(inputFile);
                Console.WriteLine($"Workbook \"{inputFile}\" loaded successfully.");

                // Iterate through worksheets and OLE objects.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    OleObjectCollection oleObjects = sheet.OleObjects;
                    for (int i = 0; i < oleObjects.Count; i++)
                    {
                        OleObject ole = oleObjects[i];
                        // Note: Direct refresh of linked OLE objects is not available in the current API.
                        // This loop can be extended if future versions expose such functionality.
                        Console.WriteLine($"Found OLE object at index {i} on sheet \"{sheet.Name}\".");
                    }
                }

                // Save the workbook after processing.
                workbook.Save(outputFile, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved as \"{outputFile}\".");
            }
            catch (Exception ex)
            {
                // Output any errors that occur during processing.
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
