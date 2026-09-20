// Title: How to set a custom display name for linked OLE objects in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells to loop through a worksheet's OleObjects collection and assign a new Name value to each OLE object in C#. | Programmatically change the display label of embedded OLE objects and save the workbook with Aspose.Cells. | Update the OleObject.Name property for all linked OLE objects in an existing .xlsx file and write the result to a new file using C#.
// Common Searches: C# Aspose.Cells change OLE object label in existing Excel file | Set OleObject.Name for linked OLE objects with Aspose.Cells .NET | Rename all OLE objects in a worksheet using Aspose.Cells API | How to modify display name of embedded OLE objects in Excel via C# code
// Tags: Aspose.Cells rename OLE objects | set OleObject.Name C# | iterate worksheet OleObjects Aspose.Cells | save workbook after OLE name update | modify OLE object display label programmatically

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing Excel workbook, iterates over every OleObject on the first worksheet, assigns a descriptive string to each object's Name property, ensures the output directory exists, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Iterate through all OLE objects on the sheet
            foreach (OleObject ole in sheet.OleObjects)
            {
                // Set a descriptive name for each OLE object
                ole.Name = "Descriptive OLE Object";
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
