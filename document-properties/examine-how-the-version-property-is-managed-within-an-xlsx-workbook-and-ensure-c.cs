using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsVersionDemo
{
    class Program
    {
        static void Main()
        {
            // -------------------------------------------------
            // Create a new workbook (lifecycle: create)
            // -------------------------------------------------
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // Set built‑in version properties
            // -------------------------------------------------
            // Version: format "00.0000" (e.g., "12.0000")
            workbook.BuiltInDocumentProperties.Version = "12.0000";

            // DocumentVersion: a custom version string for the file
            workbook.BuiltInDocumentProperties.DocumentVersion = "1.0";

            // BuildVersion: incremental public release of the application
            workbook.Settings.BuildVersion = "2.5.0";

            // -------------------------------------------------
            // Display the set values
            // -------------------------------------------------
            Console.WriteLine("Set Version: " + workbook.BuiltInDocumentProperties.Version);
            Console.WriteLine("Set DocumentVersion: " + workbook.BuiltInDocumentProperties.DocumentVersion);
            Console.WriteLine("Set BuildVersion: " + workbook.Settings.BuildVersion);

            // -------------------------------------------------
            // Save the workbook (lifecycle: save)
            // -------------------------------------------------
            string filePath = "VersionDemo.xlsx";
            workbook.Save(filePath, SaveFormat.Xlsx);

            // -------------------------------------------------
            // Load the workbook back (lifecycle: load)
            // -------------------------------------------------
            Workbook loadedWorkbook = new Workbook(filePath);

            // -------------------------------------------------
            // Retrieve and display version properties from the loaded file
            // -------------------------------------------------
            string loadedVersion = loadedWorkbook.BuiltInDocumentProperties.Version;
            string loadedDocVersion = loadedWorkbook.BuiltInDocumentProperties.DocumentVersion;
            string loadedBuildVersion = loadedWorkbook.Settings.BuildVersion;

            Console.WriteLine("Loaded Version: " + loadedVersion);
            Console.WriteLine("Loaded DocumentVersion: " + loadedDocVersion);
            Console.WriteLine("Loaded BuildVersion: " + loadedBuildVersion);
        }
    }
}