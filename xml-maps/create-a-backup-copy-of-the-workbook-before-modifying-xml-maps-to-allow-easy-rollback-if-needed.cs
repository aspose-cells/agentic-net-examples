// Title: How to create a backup of an Excel workbook before editing XML maps with Aspose.Cells for .NET (C#)
// Common Searches: C# code to copy an existing .xlsx file using Aspose.Cells before modifying XML maps | Aspose.Cells save workbook as backup before changing XmlMap properties in .NET | Create a duplicate workbook for rollback when updating XML map name with Aspose.Cells | How to preserve the original Excel file while editing XML maps using Aspose.Cells in C# | Backup Excel workbook to a separate file then overwrite after XML map changes with Aspose.Cells
// Tags: Aspose.Cells backup workbook before XML map edit | save workbook copy Xlsx Aspose.Cells C# | modify XmlMap name Aspose.Cells | rollback Excel changes Aspose.Cells XML maps | duplicate Excel file Aspose.Cells SaveFormat

using System;
using Aspose.Cells;

// The example loads an existing workbook, creates a backup copy using SaveFormat.Xlsx, checks for XML maps, modifies the first map's name if present, and finally overwrites the original file with the updated workbook.
class WorkbookBackupExample
{
    static void Main()
    {
        // Paths for the original workbook and its backup copy
        string originalPath = "input.xlsx";
        string backupPath = "input_backup.xlsx";

        // Load the original workbook
        Workbook workbook = new Workbook(originalPath);

        // Create a backup copy before making any changes
        workbook.Save(backupPath, SaveFormat.Xlsx);

        // -------------------------------------------------
        // Begin modifications to XML maps (example placeholder)
        // -------------------------------------------------
        // Ensure there is at least one XML map to work with
        if (workbook.Worksheets.XmlMaps.Count > 0)
        {
            // Access the first XML map
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

            // Example modification: change the name of the XML map
            xmlMap.Name = "ModifiedXmlMapName";

            // Additional XML map modifications can be performed here
        }
        else
        {
            // If no XML maps exist, you could add a new one as needed
            // XmlMap newMap = workbook.Worksheets.XmlMaps.Add("NewMap", "http://example.com/schema");
        }

        // Save the workbook after modifications (overwrites the original file)
        workbook.Save(originalPath, SaveFormat.Xlsx);
    }
}
