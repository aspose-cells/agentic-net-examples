// Title: Replace code of a specific VBA module in an .xlsm workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an .xlsm file with Aspose.Cells, reads a .bas file, and replaces the code of a specified VBA module while preserving macros. | Show how to add error handling that checks for a VBA project and the existence of the target module before modifying its code using Aspose.Cells. | Create a C# example that adds a new VBA module from a .bas file to an existing .xlsm workbook with Aspose.Cells.
// Common Searches: C# Aspose.Cells replace VBA module code from external .bas file in existing xlsm workbook | how to programmatically update a macro module in an xlsm file using Aspose.Cells .NET | Aspose.Cells load macro-enabled workbook and edit VBA project modules while keeping macros | replace specific VBA module in Excel macro-enabled workbook with code from .bas using Aspose.Cells
// Tags: overwrite VBA module Aspose.Cells | modify macro in .xlsm C# | preserve VBA when loading workbook Aspose.Cells | import .bas file into VBA project Aspose.Cells | save workbook SaveFormat.Xlsm

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;   // Required for VbaProject and VbaModule types

// // Loads an existing .xlsm workbook, verifies the VBA project, replaces the code of a chosen module with the contents of an external .bas file, and saves the workbook preserving macros.
class Program
{
    static void Main()
    {
        // Paths
        string workbookPath = "input.xlsm";
        string outputPath = "output.xlsm";
        string macroSourcePath = "newMacro.bas";
        string targetModuleName = "Module1";

        // Verify input files exist
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Workbook file not found: {workbookPath}");
            return;
        }

        if (!File.Exists(macroSourcePath))
        {
            Console.WriteLine($"Macro source file not found: {macroSourcePath}");
            return;
        }

        try
        {
            // Load the workbook (format is auto‑detected, macros are preserved)
            Workbook workbook = new Workbook(workbookPath);

            // Ensure the workbook contains a VBA project
            if (workbook.VbaProject == null)
            {
                Console.WriteLine("The workbook does not contain a VBA project.");
                return;
            }

            // Retrieve the target module
            VbaModule targetModule = workbook.VbaProject.Modules[targetModuleName];
            if (targetModule == null)
            {
                Console.WriteLine($"Module '{targetModuleName}' not found in the VBA project.");
                return;
            }

            // Replace module code with new macro
            string newMacroCode = File.ReadAllText(macroSourcePath);
            targetModule.Codes = newMacroCode;

            // Save the workbook, preserving macros
            workbook.Save(outputPath, SaveFormat.Xlsm);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
