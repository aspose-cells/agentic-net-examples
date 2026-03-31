using System;
using Aspose.Cells;
using Aspose.Cells.Properties;
using Aspose.Cells.Metadata;

namespace DocumentVersionScenarios
{
    class Program
    {
        static void Main()
        {
            // ------------------------------------------------------------
            // Scenario 1: Create a new workbook and set the DocumentVersion
            // ------------------------------------------------------------
            Workbook wbCreate = new Workbook(); // create a new workbook
            wbCreate.BuiltInDocumentProperties.DocumentVersion = "1.0"; // set version
            Console.WriteLine("Scenario 1 - Set version to: " + wbCreate.BuiltInDocumentProperties.DocumentVersion);
            wbCreate.Save("Scenario1_Version_1_0.xlsx", SaveFormat.Xlsx);

            // ------------------------------------------------------------
            // Scenario 2: Load an existing workbook and read its DocumentVersion
            // ------------------------------------------------------------
            Workbook wbLoad = new Workbook("Scenario1_Version_1_0.xlsx"); // load the file created above
            string loadedVersion = wbLoad.BuiltInDocumentProperties.DocumentVersion;
            Console.WriteLine("Scenario 2 - Loaded version: " + loadedVersion);

            // ------------------------------------------------------------
            // Scenario 3: Update DocumentVersion after making changes and save with strict OOXML compliance
            // ------------------------------------------------------------
            // Modify the workbook (add some data)
            Worksheet sheet = wbLoad.Worksheets[0];
            sheet.Cells["A1"].PutValue("Updated content");
            // Update the version to reflect the change
            wbLoad.BuiltInDocumentProperties.DocumentVersion = "2.0";
            // Change compliance to ISO strict (optional, shows interaction with other settings)
            wbLoad.Settings.Compliance = OoxmlCompliance.Iso29500_2008_Strict;
            Console.WriteLine("Scenario 3 - Updated version to: " + wbLoad.BuiltInDocumentProperties.DocumentVersion);
            wbLoad.Save("Scenario3_Version_2_0_Strict.xlsx", SaveFormat.Xlsx);

            // ------------------------------------------------------------
            // Scenario 4: Use WorkbookMetadata to modify DocumentVersion without loading the full workbook
            // ------------------------------------------------------------
            // Prepare metadata options to work with document properties only
            MetadataOptions metaOptions = new MetadataOptions(MetadataType.DocumentProperties);
            // Load metadata from the file created in Scenario 3
            WorkbookMetadata meta = new WorkbookMetadata("Scenario3_Version_2_0_Strict.xlsx", metaOptions);
            // Set a new version value
            meta.BuiltInDocumentProperties.DocumentVersion = "3.0";
            Console.WriteLine("Scenario 4 - Metadata version set to: " + meta.BuiltInDocumentProperties.DocumentVersion);
            // Save the modified metadata back to a new file
            meta.Save("Scenario4_Version_3_0.xlsx");

            // ------------------------------------------------------------
            // Scenario 5: Verify the final version after all modifications
            // ------------------------------------------------------------
            Workbook wbFinal = new Workbook("Scenario4_Version_3_0.xlsx");
            Console.WriteLine("Scenario 5 - Final version in workbook: " + wbFinal.BuiltInDocumentProperties.DocumentVersion);
        }
    }
}