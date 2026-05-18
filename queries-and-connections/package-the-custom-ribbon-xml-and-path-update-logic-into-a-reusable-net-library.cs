using System;
using Aspose.Cells;

namespace CustomRibbonLibrary
{
    /// <summary>
    /// Provides reusable methods for working with custom Ribbon XML and updating the library path in a workbook.
    /// </summary>
    public static class RibbonHelper
    {
        /// <summary>
        /// Creates a new workbook, assigns the provided Ribbon XML, and saves it to the specified file.
        /// </summary>
        /// <param name="ribbonXml">The custom Ribbon XML definition.</param>
        /// <param name="outputPath">The full path where the workbook will be saved (e.g., "output.xlsm").</param>
        public static void CreateWorkbookWithRibbon(string ribbonXml, string outputPath)
        {
            // Create a new workbook (empty workbook with a default worksheet)
            Workbook workbook = new Workbook();

            // Assign the custom Ribbon XML to the workbook
            workbook.RibbonXml = ribbonXml;

            // Save the workbook; the .xlsm extension ensures macro support for Ribbon UI
            workbook.Save(outputPath);
        }

        /// <summary>
        /// Loads an existing workbook, updates the Ribbon XML, optionally updates the library path used by external references,
        /// and saves the modified workbook to a new file.
        /// </summary>
        /// <param name="inputPath">Path to the existing workbook to load.</param>
        /// <param name="ribbonXml">The new Ribbon XML to assign. Pass null to leave unchanged.</param>
        /// <param name="newLibraryPath">The new library path for external formula references. Pass null to leave unchanged.</param>
        /// <param name="outputPath">Path where the modified workbook will be saved.</param>
        public static void UpdateWorkbook(string inputPath, string ribbonXml, string newLibraryPath, string outputPath)
        {
            // Load the existing workbook from disk
            Workbook workbook = new Workbook(inputPath);

            // Update Ribbon XML if a value is provided
            if (!string.IsNullOrEmpty(ribbonXml))
            {
                workbook.RibbonXml = ribbonXml;
            }

            // Update the library path used by external references if a value is provided
            if (!string.IsNullOrEmpty(newLibraryPath))
            {
                CellsHelper.LibraryPath = newLibraryPath;
            }

            // Save the modified workbook
            workbook.Save(outputPath);
        }
    }

    // Example usage of the library (can be removed or placed in a separate test project)
    class Program
    {
        static void Main()
        {
            // Sample Ribbon XML
            string sampleRibbonXml =
                "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
                "  <ribbon>" +
                "    <tabs>" +
                "      <tab id=\"customTab\" label=\"My Tab\">" +
                "        <group id=\"customGroup\" label=\"My Group\">" +
                "          <button id=\"customButton\" label=\"My Button\" size=\"large\" />" +
                "        </group>" +
                "      </tab>" +
                "    </tabs>" +
                "  </ribbon>" +
                "</customUI>";

            // Create a new workbook with the custom Ribbon and save it
            RibbonHelper.CreateWorkbookWithRibbon(sampleRibbonXml, "RibbonWorkbook.xlsm");

            // Update an existing workbook: change Ribbon XML and library path, then save as a new file
            RibbonHelper.UpdateWorkbook(
                inputPath: "RibbonWorkbook.xlsm",
                ribbonXml: sampleRibbonXml,               // reuse same XML or provide a different one
                newLibraryPath: @"C:\MyCustomLibPath",    // new path for external references
                outputPath: "RibbonWorkbook_Updated.xlsm");

            Console.WriteLine("Operations completed.");
        }
    }
}