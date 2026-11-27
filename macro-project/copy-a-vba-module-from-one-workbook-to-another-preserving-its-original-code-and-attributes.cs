// Title: Copy a VBA module from one .xlsm workbook to another using Aspose.Cells for .NET (C#)
// Common Searches: aspnet copy vba module from source xlsm to destination xlsm using Aspose.Cells C# | how to preserve VBA module attributes when moving between workbooks with Aspose.Cells | C# code to duplicate a VBA module in another workbook using Aspose.Cells | example of adding an existing VBA module to a different workbook with Aspose.Cells | transfer VBA code from one Excel file to another programmatically in .NET
// Tags: Aspose.Cells copy VBA module C# | transfer VBA module between xlsm workbooks | preserve VBA code attributes Aspose.Cells | add VBA module to destination workbook .NET | VbaProject module cloning Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

// The example loads a source and a destination .xlsm workbook, ensures both contain VBA projects, retrieves a specified VBA module from the source, adds a new module with the same type and name to the destination, copies the module's code (including attributes), and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string srcPath = "SourceWorkbook.xlsm";
            const string destPath = "DestinationWorkbook.xlsm";
            const string outputPath = "DestinationWorkbook_WithModule.xlsm";
            const string moduleName = "Module1";

            // Verify source and destination files exist
            if (!File.Exists(srcPath))
                throw new FileNotFoundException($"Source workbook not found: {srcPath}");
            if (!File.Exists(destPath))
                throw new FileNotFoundException($"Destination workbook not found: {destPath}");

            // Load workbooks
            Workbook srcWorkbook = new Workbook(srcPath);
            Workbook destWorkbook = new Workbook(destPath);

            // Access VBA projects
            VbaProject srcVbaProject = srcWorkbook.VbaProject;
            VbaProject destVbaProject = destWorkbook.VbaProject;

            if (srcVbaProject == null)
                throw new InvalidOperationException("Source workbook does not contain a VBA project.");
            if (destVbaProject == null)
                throw new InvalidOperationException("Destination workbook does not contain a VBA project.");

            // Retrieve the source module
            VbaModule srcModule = srcVbaProject.Modules[moduleName];
            if (srcModule == null)
                throw new InvalidOperationException($"Module '{moduleName}' not found in source workbook.");

            // Add a new module to the destination workbook (returns the index of the new module)
            int newModuleIndex = destVbaProject.Modules.Add(srcModule.Type, srcModule.Name);
            VbaModule destModule = destVbaProject.Modules[newModuleIndex];

            // Copy the VBA code
            destModule.Codes = srcModule.Codes;

            // Save the updated destination workbook
            destWorkbook.Save(outputPath);
            Console.WriteLine($"Module '{moduleName}' copied successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
