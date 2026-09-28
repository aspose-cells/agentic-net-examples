// Title: Rename a VBA module to DataProcessor in an Excel workbook with Aspose.Cells for .NET
// AI Prompts: Use Aspose.Cells in C# to set the Name property of the first VBA module to 'DataProcessor' and save the workbook. | Programmatically change a VBA module's identifier in an existing .xlsx file using the VbaProject API of Aspose.Cells. | Update the VBA module name before persisting the workbook with Aspose.Cells for .NET.
// Common Searches: C# Aspose.Cells change VBA module name in existing Excel file | how to set VBA module Name property using Aspose.Cells .NET API | rename first VBA module before saving workbook with Aspose.Cells | modify VBA project module identifier programmatically in C# | Aspose.Cells example for updating VBA module name
// Tags: rename VBA module Aspose.Cells .NET | set VbaProject module Name property C# | update VBA module identifier in Excel workbook | Aspose.Cells VBA project manipulation | save workbook after VBA changes .NET

using Aspose.Cells;
using Aspose.Cells.Vba;
using System;
using System.IO;

// The program loads an Excel file, accesses its VBA project, renames the first VBA module to "DataProcessor" using Aspose.Cells for .NET, and saves the modified workbook.
class RenameVbaModule
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Ensure there is at least one module to rename
            if (vbaProject != null && vbaProject.Modules.Count > 0)
            {
                // Rename the first module (or specify the desired index)
                vbaProject.Modules[0].Name = "DataProcessor";
                Console.WriteLine("VBA module renamed successfully.");
            }
            else
            {
                Console.WriteLine("No VBA modules found in the workbook.");
            }

            // Save the workbook with the renamed VBA module
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
