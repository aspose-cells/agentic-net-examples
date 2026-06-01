using System;
using System.IO;
using Aspose.Cells;

class ExtractVbaDescription
{
    static void Main()
    {
        // Path to the macro-enabled workbook
        string filePath = "sample_with_macro.xlsm";

        // Verify that the file exists before attempting to load it
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {Path.GetFullPath(filePath)}");
            return;
        }

        try
        {
            // Load the workbook from the file
            Workbook workbook = new Workbook(filePath);

            // Check if the workbook contains a VBA project
            if (workbook.HasMacro && workbook.VbaProject != null)
            {
                // Use the VBA project's Name property as its description
                string vbaDescription = workbook.VbaProject.Name;
                Console.WriteLine("VBA Project Description (Name): " + vbaDescription);
            }
            else
            {
                Console.WriteLine("The workbook does not contain a VBA project.");
            }
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}