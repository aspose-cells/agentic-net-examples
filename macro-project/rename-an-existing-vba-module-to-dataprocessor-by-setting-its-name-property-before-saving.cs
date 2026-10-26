// Title: Rename a VBA module to DataProcessor in an Excel workbook using Aspose.Cells for .NET
// Common Searches: how to change the name of a VBA module with Aspose.Cells C# | set VbaModule.Name to rename VBA module before saving workbook | Aspose.Cells rename first VBA module to DataProcessor example | C# code to rename VBA module in an existing .xlsx file using Aspose.Cells | update VBA module name in Excel file programmatically with Aspose.Cells
// Tags: Aspose.Cells VBA module renaming C# | VbaModule Name property usage | programmatic VBA module rename .NET | Excel workbook VBA project module update | C# Aspose.Cells change VBA module identifier

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;   // Required for VbaProject and VbaModule classes

// The example loads an existing .xlsx workbook, accesses its VBA project, sets the Name property of the first VbaModule to "DataProcessor", and saves the workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the VBA project (if any)
            VbaProject vbaProject = workbook.VbaProject;
            if (vbaProject != null && vbaProject.Modules.Count > 0)
            {
                // Rename the first VBA module to "DataProcessor"
                VbaModule vbaModule = vbaProject.Modules[0];
                vbaModule.Name = "DataProcessor";
            }

            // Save the workbook with the updated VBA module name
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
