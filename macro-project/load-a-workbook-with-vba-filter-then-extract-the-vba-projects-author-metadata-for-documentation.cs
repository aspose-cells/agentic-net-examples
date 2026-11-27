// Title: Read VBA project author metadata from an .xls workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# read VBA project author from .xls file | How to get VBA project author information with Aspose.Cells .NET | Check if workbook contains a VBA project and retrieve the author using Aspose.Cells | C# extract VBA project metadata (author) from Excel with Aspose.Cells | Handle missing VBA project when reading author property in Aspose.Cells
// Tags: read VBA author Aspose.Cells .NET | load workbook with VBA project C# | extract VBA project metadata Aspose.Cells | detect absent VBA project Aspose.Cells | dynamic VbaProjectInfo access C#

using System;
using System.IO;
using Aspose.Cells;

// The example loads an .xls workbook with Aspose.Cells, verifies the presence of a VBA project, and safely reads the VbaProjectInfo.Author property using dynamic binding. It outputs the author name or indicates it is unknown, while handling missing files, absent VBA projects, and runtime exceptions.
class Program
{
    static void Main()
    {
        const string filePath = "input.xls";

        // Ensure the input file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Load the workbook; VBA project is loaded automatically if present
            Workbook workbook = new Workbook(filePath);

            // Verify that a VBA project exists
            if (workbook.VbaProject != null)
            {
                // Use dynamic to safely access VbaProjectInfo (avoids compile‑time binding issues)
                dynamic vbaProject = workbook.VbaProject;
                string author = null;

                try
                {
                    author = vbaProject.VbaProjectInfo?.Author as string;
                }
                catch
                {
                    // Property not available in this version; author remains null
                }

                Console.WriteLine("VBA Project Author: " + (string.IsNullOrEmpty(author) ? "Unknown" : author));
            }
            else
            {
                Console.WriteLine("No VBA project found in the workbook.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing workbook: {ex.Message}");
        }
    }
}
