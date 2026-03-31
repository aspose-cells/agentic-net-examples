using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsEssentials
{
    class Program
    {
        static void Main()
        {
            // ------------------------------------------------------------
            // 1. License – must be applied before using most Aspose.Cells APIs.
            //    If the license is missing or invalid the library works in
            //    evaluation mode (watermarks, feature limits, etc.).
            // ------------------------------------------------------------
            try
            {
                var license = new License();
                // Provide the full path to your license file or use a stream.
                license.SetLicense("Aspose.Cells.lic");
                Console.WriteLine("License loaded successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("License could not be loaded: " + ex.Message);
            }

            // ------------------------------------------------------------
            // 2. Cloud platform flag – set to true when running in Azure,
            //    AWS Lambda, Google Cloud, etc. This influences internal
            //    path handling and temporary file locations.
            // ------------------------------------------------------------
            CellsHelper.IsCloudPlatform = true; // change to false for on‑premises

            // ------------------------------------------------------------
            // 3. Startup and library paths – required for external formula
            //    references (e.g., custom functions stored in add‑ins).
            // ------------------------------------------------------------
            CellsHelper.StartupPath = Path.Combine(Environment.CurrentDirectory, "AsposeStartup");
            CellsHelper.LibraryPath = Path.Combine(Environment.CurrentDirectory, "AsposeLibrary");

            // ------------------------------------------------------------
            // 4. Excel restriction handling – Excel imposes limits (e.g.,
            //    max string length 32,767). Setting this to false allows
            //    you to store values that exceed those limits, but you
            //    cannot save the workbook as a native Excel file afterwards.
            // ------------------------------------------------------------
            var workbook = new Workbook();
            workbook.Settings.CheckExcelRestriction = false;

            // Example: store a string longer than Excel's normal limit.
            string longString = new string('A', 40000);
            workbook.Worksheets[0].Cells["A1"].PutValue(longString);
            Console.WriteLine("Inserted long string of length: " + longString.Length);

            // ------------------------------------------------------------
            // 5. Access cache – improves performance for read‑only bulk
            //    operations. The data accessed must remain unchanged while
            //    the cache is active; otherwise an exception may be thrown.
            // ------------------------------------------------------------
            workbook.StartAccessCache(AccessCacheOptions.PositionAndSize | AccessCacheOptions.CellsData);
            // Perform read‑only actions here.
            string retrieved = workbook.Worksheets[0].Cells["A1"].StringValue;
            Console.WriteLine("Retrieved string length while cache active: " + retrieved.Length);
            workbook.CloseAccessCache(AccessCacheOptions.PositionAndSize | AccessCacheOptions.CellsData);

            // ------------------------------------------------------------
            // 6. Save options – when saving, propagate the restriction flag
            //    if you disabled it earlier. Otherwise saving to native Excel
            //    format will raise an exception.
            // ------------------------------------------------------------
            var saveOptions = new OoxmlSaveOptions
            {
                CheckExcelRestriction = false // match workbook setting
            };

            string outputFile = Path.Combine(Environment.CurrentDirectory, "LongStringDemo.xlsx");
            workbook.Save(outputFile, saveOptions);
            Console.WriteLine("Workbook saved to: " + outputFile);
        }
    }
}