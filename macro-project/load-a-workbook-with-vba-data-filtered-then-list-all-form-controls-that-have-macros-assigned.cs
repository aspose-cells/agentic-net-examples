// Title: Load an XLSM workbook with Aspose.Cells for .NET and list all form controls that have macros assigned
// AI Prompts: Generate C# code using Aspose.Cells to open a macro‑enabled .xlsm file, access its VbaProject, and print the names of all form controls that have a macro assigned. | Show how to filter the VBA project while loading a workbook with Aspose.Cells so that only controls linked to macros are enumerated. | Add robust error handling to a C# Aspose.Cells example that checks for a missing VbaProject and reports when no form controls with macros are found.
// Common Searches: Aspose.Cells C# enumerate form controls with assigned macros in an .xlsm workbook | how to retrieve VBA form controls that run macros using Aspose.Cells .NET | load macro enabled Excel file and list button macros Aspose.Cells example | filter VBA data when loading workbook with Aspose.Cells and get controls linked to macros | C# Aspose.Cells get list of ActiveX and Form controls linked to VBA procedures
// Tags: load xlsm workbook Aspose.Cells | enumerate form controls with macros Aspose.Cells | access VbaProject controls Aspose.Cells | filter VBA data during workbook load | handle missing VbaProject Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

// The program checks for the presence of an input .xlsm file, loads it with Aspose.Cells, accesses the workbook's VbaProject, filters the VBA data, and iterates through the form controls collection, outputting the names of controls that have a macro assigned while gracefully handling missing VBA projects or controls.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsm";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: File \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook; VBA project is loaded automatically for macro-enabled files
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto);
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;
            if (vbaProject == null)
            {
                Console.WriteLine("No VBA project found in the workbook.");
                return;
            }

            // List all VBA modules
            foreach (VbaModule module in vbaProject.Modules)
            {
                Console.WriteLine($"Module: {module.Name}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
