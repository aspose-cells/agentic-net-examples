// Title: Build a .NET class library to create and update Excel workbooks with custom Ribbon XML using Aspose.Cells
// AI Prompts: Write a C# static method that creates a new Workbook, assigns a RibbonXml string, optionally sets CellsHelper.AltStartPath and CellsHelper.LibraryPath, ensures the output folder exists, and saves the file with Aspose.Cells. | Implement a C# static method that loads an existing .xlsx file, replaces its RibbonXml, updates optional AltStartPath/LibraryPath, creates the destination directory if missing, and saves the modified workbook. | Generate a sample console program that demonstrates calling the library methods to create a workbook with a custom ribbon and to update the ribbon of an existing workbook.
// Common Searches: how to add custom ribbon UI to an Excel file with Aspose.Cells C# | Aspose.Cells replace RibbonXml in existing workbook programmatically | C# library method for setting AltStartPath and LibraryPath in Aspose.Cells | create reusable .NET component for Excel ribbon customization using Aspose.Cells | save Aspose.Cells workbook to a new folder that may not exist
// Tags: Aspose.Cells custom ribbon xml integration | C# static library for workbook ribbon manipulation | set AltStartPath CellsHelper Aspose.Cells | replace RibbonXml existing workbook Aspose.Cells | automatic output directory creation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace CustomRibbonLibrary
{
    // Provides a reusable .NET class library with two static methods: one to generate a new Excel workbook and embed custom Ribbon XML, and another to load an existing workbook, replace its RibbonXml, optionally configure CellsHelper.AltStartPath and CellsHelper.LibraryPath, ensure the target folder exists, and save the result using Aspose.Cells.
    public static class RibbonUtility
    {
        /// <param name="outputPath">File path where the workbook will be saved.</param>
        /// <param name="ribbonXml">Custom Ribbon XML string.</param>
        /// <param name="altStartPath">Optional alternate startup path for external references.</param>
        /// <param name="libraryPath">Optional library path for external references.</param>
        public static void CreateWorkbookWithRibbon(string outputPath, string ribbonXml,
            string? altStartPath = null, string? libraryPath = null)
        {
            try
            {
                // Create a new workbook.
                Workbook workbook = new Workbook();

                // Set the custom Ribbon XML.
                workbook.RibbonXml = ribbonXml;

                // Update alternate startup path if supplied.
                if (!string.IsNullOrEmpty(altStartPath))
                {
                    CellsHelper.AltStartPath = altStartPath;
                }

                // Update library path if supplied.
                if (!string.IsNullOrEmpty(libraryPath))
                {
                    CellsHelper.LibraryPath = libraryPath;
                }

                // Ensure the directory exists.
                string? directory = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // Save the workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook created and saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating workbook: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Loads an existing workbook, replaces its Ribbon XML, optionally updates path settings, and saves it.
        /// </summary>
        /// <param name="existingPath">Path to the existing workbook to load.</param>
        /// <param name="outputPath">Path where the modified workbook will be saved.</param>
        /// <param name="newRibbonXml">New Ribbon XML to apply.</param>
        /// <param name="altStartPath">Optional alternate startup path for external references.</param>
        /// <param name="libraryPath">Optional library path for external references.</param>
        public static void UpdateWorkbookRibbon(string existingPath, string outputPath,
            string newRibbonXml, string? altStartPath = null, string? libraryPath = null)
        {
            try
            {
                // Verify the source file exists.
                if (!File.Exists(existingPath))
                {
                    throw new FileNotFoundException($"The workbook '{existingPath}' does not exist.");
                }

                // Load the workbook from file.
                Workbook workbook = new Workbook(existingPath);

                // Replace the Ribbon XML.
                workbook.RibbonXml = newRibbonXml;

                // Update paths if provided.
                if (!string.IsNullOrEmpty(altStartPath))
                {
                    CellsHelper.AltStartPath = altStartPath;
                }

                if (!string.IsNullOrEmpty(libraryPath))
                {
                    CellsHelper.LibraryPath = libraryPath;
                }

                // Ensure the output directory exists.
                string? outDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outDir) && !Directory.Exists(outDir))
                {
                    Directory.CreateDirectory(outDir);
                }

                // Save the modified workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook '{existingPath}' updated and saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error updating workbook: {ex.Message}");
                throw;
            }
        }
    }

    internal class Program
    {
        private static void Main(string[] args)
        {
            // Sample Ribbon XML (replace with actual XML as needed).
            string sampleRibbonXml = @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon>
    <tabs>
      <tab id='customTab' label='Custom Tab'>
        <group id='customGroup' label='Custom Group'>
          <button id='customButton' label='Click Me' size='large' onAction='OnButtonClick' />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";

            // Paths for demonstration.
            string newWorkbookPath = Path.Combine(Environment.CurrentDirectory, "Output", "NewWorkbook.xlsx");
            string existingWorkbookPath = Path.Combine(Environment.CurrentDirectory, "Input", "ExistingWorkbook.xlsx");
            string updatedWorkbookPath = Path.Combine(Environment.CurrentDirectory, "Output", "UpdatedWorkbook.xlsx");

            // Create a new workbook with Ribbon XML.
            RibbonUtility.CreateWorkbookWithRibbon(newWorkbookPath, sampleRibbonXml);

            // Update an existing workbook's Ribbon XML if the source file exists.
            if (File.Exists(existingWorkbookPath))
            {
                RibbonUtility.UpdateWorkbookRibbon(existingWorkbookPath, updatedWorkbookPath, sampleRibbonXml);
            }
            else
            {
                Console.WriteLine($"Source workbook '{existingWorkbookPath}' not found. Skipping update.");
            }
        }
    }
}
