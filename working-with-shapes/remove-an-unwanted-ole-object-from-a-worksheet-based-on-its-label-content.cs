// Title: Remove an OLE object with a specific label from an Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells, finds an OLE object whose Name matches a given label, deletes it, and saves the file. | Show how to iterate the Worksheet.OleObjects collection in reverse order and remove the matching OLE object by its label safely. | Create a reusable method RemoveOleObject(string inputPath, string label, string outputPath) that purges the unwanted OLE object using Aspose.Cells.
// Common Searches: Aspose.Cells C# remove OLE object by name from worksheet | how to delete embedded OLE object in Excel using Aspose.Cells .NET | iterate OleObjectCollection and delete specific object Aspose.Cells | C# code to purge unwanted OLE object from Excel file with Aspose.Cells | remove OLE object with label UnwantedLabel using Aspose.Cells
// Tags: remove OLE object Aspose.Cells | OleObjectCollection iteration C# | delete embedded OLE object Excel .NET | filter OLE objects by name Aspose.Cells | save workbook after OLE removal Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// Loads an Excel workbook, searches the first worksheet for an OLE object whose Name equals a specified label, removes that object from the OleObjectCollection, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index or name as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the label (name) of the OLE object that should be removed
            string unwantedLabel = "UnwantedLabel";

            // Get the collection of OLE objects on the worksheet
            OleObjectCollection oleObjects = worksheet.OleObjects;

            // Iterate backwards so removal does not affect the loop index
            for (int i = oleObjects.Count - 1; i >= 0; i--)
            {
                OleObject ole = oleObjects[i];

                // Compare the OLE object's name (label) with the target label
                if (ole.Name == unwantedLabel)
                {
                    // Remove the matching OLE object from the worksheet
                    oleObjects.RemoveAt(i);
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
