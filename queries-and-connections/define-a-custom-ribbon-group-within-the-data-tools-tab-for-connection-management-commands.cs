// Title: Inject a custom Connection Management ribbon group into the Data Tools tab of an Excel workbook using Aspose.Cells for .NET
// Common Searches: asp.net add custom ribbon group to Data Tools tab with Aspose.Cells | c# inject customUI.xml into an xlsx package using ZipArchive for Excel ribbon customization | create connection management buttons in Excel ribbon via Aspose.Cells API | how to add Refresh and Edit Connections buttons to the Data Tools ribbon in C# | verify workbook loads after custom ribbon injection with Aspose.Cells
// Tags: Aspose.Cells custom ribbon injection | customUI XML injection into XLSX | C# ZipArchive add custom ribbon to Excel workbook | Data Tools tab custom group Aspose.Cells | connection management ribbon buttons Excel C#

using System;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;

namespace CustomRibbonDemo
{
    // The example creates a simple workbook with Aspose.Cells, saves it as an .xlsx file, then opens the package with ZipArchive to inject a customUI XML definition that adds a "Connection Management" group containing Refresh and Edit buttons under the built‑in Data Tools tab, and finally reloads the workbook to confirm the custom ribbon does not affect loading.
    class Program
    {
        // Ribbon XML that adds a custom group called "Connection Management"
        // under the built‑in "Data Tools" tab (TabDataTools).
        private const string RibbonXml = @"<?xml version=""1.0"" encoding=""UTF-8""?>
<customUI xmlns=""http://schemas.microsoft.com/office/2006/01/customui"">
  <ribbon>
    <tabs>
      <tab idMso=""TabDataTools"">
        <group id=""CustomConnectionGroup"" label=""Connection Management"">
          <button id=""RefreshConnections""
                  label=""Refresh Connections""
                  size=""large""
                  imageMso=""RefreshAll""
                  onAction=""OnRefreshConnections"" />
          <button id=""EditConnections""
                  label=""Edit Connections""
                  size=""large""
                  imageMso=""DataConnectionProperties""
                  onAction=""OnEditConnections"" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";

        static void Main(string[] args)
        {
            // Path for the temporary workbook that will receive the custom ribbon.
            string workbookPath = Path.Combine(Directory.GetCurrentDirectory(), "SampleWithCustomRibbon.xlsx");

            try
            {
                // ------------------------------------------------------------
                // 1. Create a simple workbook using Aspose.Cells.
                // ------------------------------------------------------------
                Workbook workbook = new Workbook();

                // Add some sample data – this part is optional and only for demo.
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;
                cells["A1"].PutValue("Item");
                cells["B1"].PutValue("Price");
                cells["A2"].PutValue("Apple");
                cells["B2"].PutValue(2.5);
                cells["A3"].PutValue("Orange");
                cells["B3"].PutValue(1.8);
                cells["A4"].PutValue("Total:");
                cells["B4"].Formula = "=SUM(B2:B3)";

                // Save the workbook – we will later inject the ribbon XML into this file.
                workbook.Save(workbookPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to: {workbookPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating or saving workbook: {ex.Message}");
                return;
            }

            // ------------------------------------------------------------
            // 2. Inject the custom ribbon XML into the .xlsx package.
            // ------------------------------------------------------------
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine("Workbook file not found. Aborting ribbon injection.");
                return;
            }

            try
            {
                // Open the existing .xlsx package for update using ZipArchive.
                using (FileStream fs = new FileStream(workbookPath, FileMode.Open, FileAccess.ReadWrite))
                using (ZipArchive archive = new ZipArchive(fs, ZipArchiveMode.Update, leaveOpen: true))
                {
                    const string entryPath = "customUI/customUI.xml";

                    // Remove existing entry if it exists.
                    ZipArchiveEntry existingEntry = archive.GetEntry(entryPath);
                    existingEntry?.Delete();

                    // Create a new entry and write the ribbon XML.
                    ZipArchiveEntry entry = archive.CreateEntry(entryPath, CompressionLevel.Optimal);
                    using (StreamWriter writer = new StreamWriter(entry.Open()))
                    {
                        writer.Write(RibbonXml);
                    }
                }

                Console.WriteLine("Custom ribbon XML successfully injected into the workbook.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error while injecting ribbon XML: {ex.Message}");
                return;
            }

            // ------------------------------------------------------------
            // 3. Verify that the file can still be opened by Aspose.Cells.
            // ------------------------------------------------------------
            try
            {
                Workbook verifyWorkbook = new Workbook(workbookPath);
                Console.WriteLine("Verification: Workbook loaded successfully after ribbon injection.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Verification failed: {ex.Message}");
            }

            Console.WriteLine("Done. Open the file in Excel to see the new \"Connection Management\" group under the Data Tools tab.");
        }
    }
}
