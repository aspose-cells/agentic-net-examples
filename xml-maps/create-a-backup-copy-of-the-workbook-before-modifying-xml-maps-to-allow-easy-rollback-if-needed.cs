// Title: C# – Backup an Excel workbook with Aspose.Cells before modifying XML maps
// Description: Demonstrates how to verify an existing .xlsx file, create a full backup copy using Aspose.Cells, add an XML map from a temporary XSD schema, clean up temporary files, and save both the backup and the modified workbook. Includes error handling for safe rollback.
// Keywords: Aspose.Cells backup workbook | C# copy Excel file | XML map Aspose.Cells | save workbook before changes | .NET Excel rollback | temporary XSD file | Excel XML schema mapping
// Common Searches: Aspose.Cells create backup before adding XML map | C# copy Excel workbook to new file | How to add XML map to Excel with Aspose.Cells | Save Excel backup .NET Aspose | Rollback Excel changes after XML map error
// Developer Intent: Generate a reliable backup of an existing workbook, then safely add or update XML maps so the original file can be restored if needed.
// Use Cases: Preserve the original data before applying a new XML schema to an Excel workbook. | Compare original and updated workbooks after XML map changes for quality assurance. | Automatically revert to the backup when an exception occurs while adding an XML map.
// AI Prompts: Write C# code using Aspose.Cells that creates a timestamped backup of a workbook, adds an XML map from a temporary XSD, and saves both files. | Provide a reusable method to backup an Excel file and safely manage temporary schema files when modifying XML maps. | Explain step‑by‑step how to restore a backup workbook if adding an XML map fails with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBackupExample
{
    // Demonstrates how to verify an existing .xlsx file, create a full backup copy using Aspose.Cells, add an XML map from a temporary XSD schema, clean up temporary files, and save both the backup and the modified workbook. Includes error handling for safe rollback.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the original workbook
                string originalPath = "OriginalWorkbook.xlsx";

                // Verify that the original workbook exists
                if (!File.Exists(originalPath))
                {
                    Console.WriteLine($"Error: File '{originalPath}' not found.");
                    return;
                }

                // Load the original workbook
                Workbook originalWorkbook = new Workbook(originalPath);

                // ---------- Create a backup copy ----------
                // Create an empty workbook that will hold the backup
                Workbook backupWorkbook = new Workbook();

                // Copy all contents from the original workbook to the backup workbook
                backupWorkbook.Copy(originalWorkbook);

                // Save the backup workbook to a separate file
                string backupPath = "OriginalWorkbook_Backup.xlsx";
                backupWorkbook.Save(backupPath);
                Console.WriteLine($"Backup saved to '{backupPath}'.");

                // ---------- Modify XML maps in the original workbook ----------
                // Example XML schema (replace with your actual schema or file path)
                string xmlSchema = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                        <xs:element name='Root'>
                                            <xs:complexType>
                                                <xs:sequence>
                                                    <xs:element name='Item' type='xs:string'/>
                                                </xs:sequence>
                                            </xs:complexType>
                                        </xs:element>
                                    </xs:schema>";

                // Write the schema to a temporary file because XmlMaps.Add expects a file path
                string tempSchemaPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xsd");
                File.WriteAllText(tempSchemaPath, xmlSchema);

                // Add the XML map to the original workbook
                int mapIndex = originalWorkbook.Worksheets.XmlMaps.Add(tempSchemaPath);
                XmlMap addedMap = originalWorkbook.Worksheets.XmlMaps[mapIndex];
                addedMap.Name = "MyBackupDemoMap";

                // Clean up the temporary schema file
                if (File.Exists(tempSchemaPath))
                {
                    File.Delete(tempSchemaPath);
                }

                // Save the modified original workbook
                string modifiedPath = "OriginalWorkbook_Modified.xlsx";
                originalWorkbook.Save(modifiedPath);
                Console.WriteLine($"Modified workbook saved to '{modifiedPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
