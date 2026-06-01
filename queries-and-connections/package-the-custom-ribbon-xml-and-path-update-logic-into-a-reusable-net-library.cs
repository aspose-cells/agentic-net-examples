using System;
using System.IO;
using Aspose.Cells;

namespace CustomRibbonLib
{
    /// <summary>
    /// Provides helper methods to work with custom Ribbon XML and path settings for Aspose.Cells workbooks.
    /// </summary>
    public static class RibbonHelper
    {
        /// <summary>
        /// Sets the custom Ribbon XML for the given workbook.
        /// </summary>
        /// <param name="workbook">The workbook to modify.</param>
        /// <param name="ribbonXml">The Ribbon XML string.</param>
        public static void SetRibbonXml(Workbook workbook, string ribbonXml)
        {
            // Assign the custom Ribbon XML.
            workbook.RibbonXml = ribbonXml;
        }

        /// <summary>
        /// Updates the library path that Aspose.Cells uses for external references.
        /// </summary>
        /// <param name="libraryPath">The full path to the library folder.</param>
        public static void UpdateLibraryPath(string libraryPath)
        {
            // Set the lookup path for external references.
            CellsHelper.LibraryPath = libraryPath;
        }

        /// <summary>
        /// Saves the workbook to the specified file path.
        /// </summary>
        /// <param name="workbook">The workbook to save.</param>
        /// <param name="outputPath">The full file name (including extension) where the workbook will be saved.</param>
        public static void SaveWorkbook(Workbook workbook, string outputPath)
        {
            // Persist the workbook to disk.
            workbook.Save(outputPath);
        }

        /// <summary>
        /// Convenience method that creates a new workbook, applies Ribbon XML, updates the library path,
        /// and saves the result.
        /// </summary>
        /// <param name="ribbonXml">Custom Ribbon XML.</param>
        /// <param name="libraryPath">Path to external libraries.</param>
        /// <param name="outputPath">Destination file path.</param>
        public static void CreateAndSave(string ribbonXml, string libraryPath, string outputPath)
        {
            // Create a new workbook.
            Workbook wb = new Workbook();

            // Apply Ribbon XML.
            SetRibbonXml(wb, ribbonXml);

            // Update library path for external references.
            UpdateLibraryPath(libraryPath);

            // Save the workbook.
            SaveWorkbook(wb, outputPath);
        }
    }

    /// <summary>
    /// Entry point for the application.
    /// </summary>
    public static class Program
    {
        public static void Main()
        {
            try
            {
                // Sample Ribbon XML (replace with actual XML as needed).
                string ribbonXml = @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
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

                // Path to external libraries (ensure the directory exists).
                string libraryPath = @"C:\AsposeLibraries";

                if (!Directory.Exists(libraryPath))
                {
                    Console.WriteLine($"Library path does not exist: {libraryPath}");
                    return;
                }

                // Destination workbook path.
                string outputPath = Path.Combine(Environment.CurrentDirectory, "CustomRibbonWorkbook.xlsx");

                // Ensure we can write to the output directory.
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!Directory.Exists(outputDir))
                {
                    Console.WriteLine($"Output directory does not exist: {outputDir}");
                    return;
                }

                // Create workbook with custom Ribbon and save it.
                RibbonHelper.CreateAndSave(ribbonXml, libraryPath, outputPath);

                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors.
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}