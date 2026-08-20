// Title: Aspose.Cells .NET Library to Add or Update Custom Ribbon XML in XLSM Workbooks
// Description: A reusable C# library that reads a custom Ribbon XML file, assigns it to a new or existing Aspose.Cells Workbook via the RibbonXml property, sets CellsHelper.LibraryPath for relative add‑in resolution, and saves the file as a macro‑enabled XLSM workbook. Includes robust error handling for creation and update scenarios.
// Keywords: Aspose.Cells RibbonXml | custom ribbon xml .NET | C# Excel macro enabled workbook | CellsHelper LibraryPath | create XLSM with ribbon | update ribbon xml Aspose.Cells | Excel UI customization C# | reusable ribbon helper library
// Common Searches: how to embed custom ribbon xml in an Excel file using Aspose.Cells | update ribbon xml of an existing .xlsm workbook with Aspose.Cells C# | set CellsHelper.LibraryPath for ribbon add‑ins in Aspose.Cells | Aspose.Cells create workbook with custom ribbon UI | C# library to replace Excel ribbon XML programmatically
// Developer Intent: Provide a ready‑to‑use .NET helper that injects or replaces a custom Ribbon XML in an Excel workbook and ensures the library path is correctly configured for add‑in references.
// Use Cases: Generate a corporate‑branded XLSM workbook that ships with a predefined Ribbon UI. | Refresh the Ribbon UI of an existing macro‑enabled workbook after a design overhaul without modifying other content. | Automate the application of a standard Ribbon XML to multiple workbooks in a CI/CD pipeline.
// AI Prompts: Write NUnit tests for RibbonHelper.CreateWorkbookWithRibbon and RibbonHelper.UpdateWorkbookRibbon covering missing files, invalid paths, and successful execution. | Modify RibbonHelper to accept a raw Ribbon XML string instead of a file path and update the documentation accordingly. | Show how to integrate Serilog into RibbonHelper for detailed error logging while preserving the existing exception flow.

using System;
using System.IO;
using Aspose.Cells;

namespace CustomRibbonLibrary
{
    // A reusable C# library that reads a custom Ribbon XML file, assigns it to a new or existing Aspose.Cells Workbook via the RibbonXml property, sets CellsHelper.LibraryPath for relative add‑in resolution, and saves the file as a macro‑enabled XLSM workbook. Includes robust error handling for creation and update scenarios.
    public static class RibbonHelper
    {
        /// <param name="ribbonXmlPath">Full path to the custom Ribbon XML file.</param>
        /// <param name="outputPath">Full path where the workbook will be saved.</param>
        public static void CreateWorkbookWithRibbon(string ribbonXmlPath, string outputPath)
        {
            try
            {
                // Validate input paths
                if (string.IsNullOrEmpty(ribbonXmlPath))
                    throw new ArgumentException("Ribbon XML path must be provided.", nameof(ribbonXmlPath));

                if (!File.Exists(ribbonXmlPath))
                    throw new FileNotFoundException("Ribbon XML file not found.", ribbonXmlPath);

                // Read the Ribbon XML content
                string ribbonXml = File.ReadAllText(ribbonXmlPath);

                // Create a new workbook (Aspose.Cells Workbook)
                Workbook workbook = new Workbook();

                // Assign the Ribbon XML to the workbook
                workbook.RibbonXml = ribbonXml;

                // Update the library path so that any external references (e.g., add‑ins) can resolve relative paths
                CellsHelper.LibraryPath = Path.GetDirectoryName(ribbonXmlPath);

                // Save the workbook in macro‑enabled format to preserve the Ribbon UI
                workbook.Save(outputPath, SaveFormat.Xlsm);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error creating workbook with ribbon: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Loads an existing workbook, replaces its Ribbon XML with the content from the specified file,
        /// updates the library path, and saves the workbook back.
        /// </summary>
        /// <param name="workbookPath">Path to the existing workbook.</param>
        /// <param name="ribbonXmlPath">Path to the new Ribbon XML file.</param>
        public static void UpdateWorkbookRibbon(string workbookPath, string ribbonXmlPath)
        {
            try
            {
                // Validate inputs
                if (!File.Exists(workbookPath))
                    throw new FileNotFoundException("Workbook file not found.", workbookPath);

                if (!File.Exists(ribbonXmlPath))
                    throw new FileNotFoundException("Ribbon XML file not found.", ribbonXmlPath);

                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Read new Ribbon XML
                string ribbonXml = File.ReadAllText(ribbonXmlPath);

                // Set the RibbonXml property
                workbook.RibbonXml = ribbonXml;

                // Update library path to the directory of the Ribbon XML
                CellsHelper.LibraryPath = Path.GetDirectoryName(ribbonXmlPath);

                // Save changes (preserve macro‑enabled format if applicable)
                workbook.Save(workbookPath, SaveFormat.Xlsm);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error updating workbook ribbon: {ex.Message}");
                throw;
            }
        }
    }

    // Simple entry point for demonstration / testing purposes
    internal class Program
    {
        private static void Main(string[] args)
        {
            try
            {
                // Example usage:
                // args[0] = "create" or "update"
                // For "create": args[1] = ribbonXmlPath, args[2] = outputPath
                // For "update": args[1] = workbookPath, args[2] = ribbonXmlPath

                if (args.Length < 3)
                {
                    Console.WriteLine("Usage:");
                    Console.WriteLine("  create <RibbonXmlPath> <OutputWorkbookPath>");
                    Console.WriteLine("  update <WorkbookPath> <RibbonXmlPath>");
                    return;
                }

                string command = args[0].ToLowerInvariant();

                if (command == "create")
                {
                    string ribbonXmlPath = args[1];
                    string outputPath = args[2];
                    RibbonHelper.CreateWorkbookWithRibbon(ribbonXmlPath, outputPath);
                    Console.WriteLine($"Workbook created at: {outputPath}");
                }
                else if (command == "update")
                {
                    string workbookPath = args[1];
                    string ribbonXmlPath = args[2];
                    RibbonHelper.UpdateWorkbookRibbon(workbookPath, ribbonXmlPath);
                    Console.WriteLine($"Workbook updated: {workbookPath}");
                }
                else
                {
                    Console.WriteLine("Invalid command. Use 'create' or 'update'.");
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
