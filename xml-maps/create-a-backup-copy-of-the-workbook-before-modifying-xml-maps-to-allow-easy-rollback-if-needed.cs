// Title: Backing up an Excel workbook before editing XML maps with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells backup workbook before adding or renaming XML maps | How to create a rollback copy of an Excel file when working with XML maps in Aspose.Cells | Save a backup of a .xlsx file using Aspose.Cells before modifying XmlMapCollection | Aspose.Cells example to copy workbook, modify XML maps, and restore on error | C# code to backup and restore Excel workbook while updating XML maps with Aspose.Cells
// Tags: Aspose.Cells create copy of workbook before XML map edit | C# rename XML map using Aspose.Cells | Add XML map from XSD schema with Aspose.Cells | Export XML map to file using Aspose.Cells C# | Aspose.Cells restore workbook from copy after XML map failure

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsXmlMapBackupDemo
{
    // The sample determines the source workbook path, creates a duplicate copy, loads the workbook with Aspose.Cells, and saves the copy as a safety net. It then accesses the XmlMapCollection to rename an existing map, add a new map from an XSD schema, and export the new map. If any operation fails, the original workbook is restored from the duplicate. Finally, the modified workbook is saved as a separate file.
    class Program
    {
        /// <param name="args">
        /// args[0] (optional) – Path to the source workbook. If omitted, "SourceWorkbook.xlsx" in the executable folder is used.
        /// </param>
        static void Main(string[] args)
        {
            // -----------------------------------------------------------------
            // 1. Determine source workbook path
            // -----------------------------------------------------------------
            string sourcePath = args.Length > 0
                ? args[0]
                : Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SourceWorkbook.xlsx");

            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Error: Source workbook not found at '{sourcePath}'.");
                return;
            }

            // -----------------------------------------------------------------
            // 2. Prepare backup file path
            // -----------------------------------------------------------------
            string backupPath = Path.Combine(
                Path.GetDirectoryName(sourcePath) ?? string.Empty,
                Path.GetFileNameWithoutExtension(sourcePath) + "_Backup" + Path.GetExtension(sourcePath));

            // -----------------------------------------------------------------
            // 3. Load the workbook and create a backup copy
            // -----------------------------------------------------------------
            Workbook workbook = null;
            try
            {
                workbook = new Workbook(sourcePath);
                // Save a physical backup copy before any modifications
                workbook.Save(backupPath);
                Console.WriteLine($"Backup created successfully at '{backupPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load or backup the workbook: {ex.Message}");
                return;
            }

            // -----------------------------------------------------------------
            // 4. Modify XML maps
            // -----------------------------------------------------------------
            try
            {
                // Access the collection of XML maps associated with the workbook
                XmlMapCollection xmlMaps = workbook.Worksheets.XmlMaps;

                // Example 1: Rename the first existing XML map (if any)
                if (xmlMaps.Count > 0)
                {
                    XmlMap firstMap = xmlMaps[0];
                    string originalName = firstMap.Name;
                    firstMap.Name = "Renamed_" + originalName;
                    Console.WriteLine($"Renamed XML map '{originalName}' to '{firstMap.Name}'.");
                }
                else
                {
                    Console.WriteLine("No existing XML maps found. Skipping rename step.");
                }

                // Example 2: Add a new XML map from an XSD schema string
                string xmlSchema = @"
<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
  <xs:element name='Products'>
    <xs:complexType>
      <xs:sequence>
        <xs:element name='Product' maxOccurs='unbounded'>
          <xs:complexType>
            <xs:sequence>
              <xs:element name='ID' type='xs:int'/>
              <xs:element name='Name' type='xs:string'/>
              <xs:element name='Price' type='xs:decimal'/>
            </xs:sequence>
          </xs:complexType>
        </xs:element>
      </xs:sequence>
    </xs:complexType>
  </xs:element>
</xs:schema>";

                int newMapIndex = xmlMaps.Add(xmlSchema);
                XmlMap newMap = xmlMaps[newMapIndex];
                newMap.Name = "ProductsMap";
                Console.WriteLine($"Added new XML map with name '{newMap.Name}' at index {newMapIndex}.");

                // Example 3: Export the newly added XML map to verify it works (optional)
                string exportPath = Path.Combine(
                    Path.GetDirectoryName(sourcePath) ?? string.Empty,
                    "ExportedProducts.xml");
                workbook.ExportXml(newMap.Name, exportPath);
                Console.WriteLine($"Exported XML map '{newMap.Name}' to '{exportPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error while modifying XML maps: {ex.Message}");
                // -----------------------------------------------------------------
                // 5. Rollback: restore the original workbook from backup
                // -----------------------------------------------------------------
                try
                {
                    File.Copy(backupPath, sourcePath, overwrite: true);
                    Console.WriteLine("Rollback successful: original workbook restored from backup.");
                }
                catch (Exception rollbackEx)
                {
                    Console.WriteLine($"Rollback failed: {rollbackEx.Message}");
                }
                return;
            }

            // -----------------------------------------------------------------
            // 6. Save the modified workbook
            // -----------------------------------------------------------------
            string modifiedPath = Path.Combine(
                Path.GetDirectoryName(sourcePath) ?? string.Empty,
                Path.GetFileNameWithoutExtension(sourcePath) + "_Modified" + Path.GetExtension(sourcePath));

            try
            {
                workbook.Save(modifiedPath);
                Console.WriteLine($"Modified workbook saved successfully at '{modifiedPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save modified workbook: {ex.Message}");
                // Attempt rollback again if saving fails
                try
                {
                    File.Copy(backupPath, sourcePath, overwrite: true);
                    Console.WriteLine("Rollback successful: original workbook restored from backup.");
                }
                catch (Exception rollbackEx)
                {
                    Console.WriteLine($"Rollback failed: {rollbackEx.Message}");
                }
            }
        }
    }
}
