using System;
using System.IO;
using Aspose.Cells;

namespace RibbonUtilityApp
{
    public static class RibbonUtility
    {
        // Creates a new workbook, assigns custom ribbon XML, and saves it as a macro‑enabled file.
        public static void CreateWorkbookWithRibbon(string ribbonXml, string outputFile)
        {
            try
            {
                Workbook workbook = new Workbook();
                workbook.RibbonXml = ribbonXml;
                workbook.Save(outputFile, SaveFormat.Xlsm);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error creating workbook: {ex.Message}");
                throw;
            }
        }

        // Loads an existing workbook, updates its Ribbon XML, and saves the result.
        public static void UpdateWorkbookRibbon(string inputFile, string ribbonXml, string outputFile)
        {
            try
            {
                if (!File.Exists(inputFile))
                    throw new FileNotFoundException("Input file not found.", inputFile);

                Workbook workbook = new Workbook(inputFile);
                workbook.RibbonXml = ribbonXml;
                workbook.Save(outputFile, SaveFormat.Xlsm);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error updating workbook: {ex.Message}");
                throw;
            }
        }

        // Sets the global library path used by Aspose.Cells for external formula references.
        public static void SetLibraryPath(string path)
        {
            try
            {
                CellsHelper.LibraryPath = path;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error setting library path: {ex.Message}");
                throw;
            }
        }
    }

    class Program
    {
        // Entry point required for compilation.
        static void Main(string[] args)
        {
            // Simple command‑line usage:
            //   create <ribbonXml> <outputFile>
            //   update <inputFile> <ribbonXml> <outputFile>
            if (args.Length == 0)
            {
                Console.WriteLine("Usage:");
                Console.WriteLine("  create <ribbonXml> <outputFile>");
                Console.WriteLine("  update <inputFile> <ribbonXml> <outputFile>");
                return;
            }

            try
            {
                string command = args[0].ToLowerInvariant();

                if (command == "create" && args.Length == 3)
                {
                    RibbonUtility.CreateWorkbookWithRibbon(args[1], args[2]);
                    Console.WriteLine("Workbook created successfully.");
                }
                else if (command == "update" && args.Length == 4)
                {
                    RibbonUtility.UpdateWorkbookRibbon(args[1], args[2], args[3]);
                    Console.WriteLine("Workbook updated successfully.");
                }
                else
                {
                    Console.WriteLine("Invalid arguments.");
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}