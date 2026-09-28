// Title: How to delete a named slicer from the first worksheet of an Excel file with Aspose.Cells for .NET
// AI Prompts: Write C# using Aspose.Cells to locate a slicer named 'Slicer1' on the first worksheet, remove it, and save the workbook as a new file. | Show a try‑catch example that verifies the source Excel file exists, deletes the specified slicer, and handles any errors with Aspose.Cells. | Provide code that iterates the worksheet's Slicers collection, removes the matching slicer, and persists the changes to output.xlsx.
// Common Searches: aspnet remove slicer by name from excel worksheet using Aspose.Cells | c# code to delete Excel slicer and save workbook with Aspose.Cells | how to programmatically find and remove a slicer in an Aspose.Cells workbook | example of cleaning up slicer connections after deletion in Aspose.Cells for .NET | remove specific slicer from first sheet Aspose.Cells C# tutorial
// Tags: Aspose.Cells delete slicer C# | remove slicer from worksheet Aspose.Cells | slicer cleanup Aspose.Cells .NET | Excel slicer removal using Aspose.Cells | C# Aspose.Cells slicer management | delete slicer connections Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing; // Required for Slicer class

// The program loads an Excel file, searches the first worksheet for a slicer named 'Slicer1', removes it if found, and saves the modified workbook to a new file while handling missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Name of the slicer to be removed
            string slicerName = "Slicer1";

            // Locate the slicer in the worksheet's slicer collection
            int slicerIndex = -1;
            for (int i = 0; i < worksheet.Slicers.Count; i++)
            {
                if (worksheet.Slicers[i].Name == slicerName)
                {
                    slicerIndex = i;
                    break;
                }
            }

            if (slicerIndex != -1)
            {
                // Remove the slicer from the worksheet
                worksheet.Slicers.RemoveAt(slicerIndex);
                Console.WriteLine($"Slicer '{slicerName}' removed.");
            }
            else
            {
                Console.WriteLine($"Slicer '{slicerName}' not found.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
