// Title: Copy worksheet page setup and set left margin to 0.5 inches with Aspose.Cells for .NET
// AI Prompts: Use Aspose.Cells to duplicate the PageSetup from Sheet1 to Sheet2 and then set the destination sheet's left margin to 0.5 inches in C#. | Copy the entire page layout of a source worksheet to another worksheet and adjust the left printing margin to half an inch using the Aspose.Cells API. | Programmatically transfer page setup between worksheets and modify the left margin value to 0.5 inches in a .NET workbook.
// Common Searches: Aspose.Cells copy page setup from one worksheet to another C# | how to change left margin to 0.5 inches after copying page setup Aspose.Cells | C# set worksheet left printing margin using Aspose.Cells PageSetup | copy page layout between sheets and adjust margins Aspose.Cells example
// Tags: worksheet page setup copy Aspose.Cells | set left margin inches Aspose.Cells | modify printing margins C# workbook | copy page layout between worksheets .NET | adjust worksheet left margin after copy

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing workbook, copies the PageSetup from a source worksheet to a destination worksheet using Aspose.Cells, sets the destination's left margin to 0.5 inches, and saves the modified file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the source worksheet (the one with the original page setup)
            Worksheet sourceSheet = workbook.Worksheets["Sheet1"]; // adjust name if needed
            if (sourceSheet == null)
            {
                Console.WriteLine("Error: Source worksheet \"Sheet1\" not found.");
                return;
            }

            // Get the destination worksheet (the one to receive the page setup)
            Worksheet destinationSheet = workbook.Worksheets["Sheet2"]; // adjust name if needed
            if (destinationSheet == null)
            {
                Console.WriteLine("Error: Destination worksheet \"Sheet2\" not found.");
                return;
            }

            // Copy the entire page setup from the source to the destination worksheet
            // Provide a CopyOptions instance as required by the API
            destinationSheet.PageSetup.Copy(sourceSheet.PageSetup, new CopyOptions());

            // Modify the left margin of the destination worksheet to 0.5 inches
            destinationSheet.PageSetup.LeftMargin = 0.5;

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
