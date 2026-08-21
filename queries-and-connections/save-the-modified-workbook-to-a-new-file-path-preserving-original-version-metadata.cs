// Title: Preserve Original Document Properties When Saving a Modified Excel Workbook with AspNet Aspose.Cells
// Description: Load an existing XLSX file, modify its content, and save it as a new workbook while retaining the source file's version, author, creation date, and custom properties. The example demonstrates using Aspose.Cells' MetadataOptions and WorkbookMetadata classes to copy document properties from the original workbook to the new file in C#.
// Keywords: Aspose.Cells preserve document properties | C# copy Excel metadata | WorkbookMetadata Aspose.Cells | MetadataOptions document properties | save modified workbook with original metadata | .NET Excel version information | retain author creation date Excel | copy custom properties Aspose.Cells
// Common Searches: how to keep original document properties when saving an edited Excel file using Aspose.Cells | copy metadata from one workbook to another Aspose.Cells .NET | preserve version information after modifying XLSX with Aspose.Cells | Aspose.Cells retain author and creation date on save | C# preserve custom Excel properties after edit
// Developer Intent: Save an edited Excel workbook while retaining all original document and version metadata.
// Use Cases: Update a template workbook and generate a new copy that maintains the original author, creation date, and custom compliance fields. | Run batch modifications on multiple Excel files while preserving each file's audit‑trail metadata for regulatory reporting. | Create versioned backups after applying business logic, ensuring the backup inherits the source file's version number and property values.
// AI Prompts: Generate C# code that loads an XLSX with Aspose.Cells, changes specific cells, and saves it to a new file while copying all original document properties using MetadataOptions and WorkbookMetadata. | Explain step‑by‑step how to transfer version and custom metadata from a source workbook to a destination workbook in Aspose.Cells for .NET, including required namespaces and object initialization.

using System;
using Aspose.Cells;
using Aspose.Cells.Metadata;

namespace AsposeCellsMetadataPreserveDemo
{
    // Load an existing XLSX file, modify its content, and save it as a new workbook while retaining the source file's version, author, creation date, and custom properties. The example demonstrates using Aspose.Cells' MetadataOptions and WorkbookMetadata classes to copy document properties from the original workbook to the new file in C#.
    public class Program
    {
        public static void Main()
        {
            // Paths for the original workbook and the new workbook
            string sourcePath = "original.xlsx";
            string destinationPath = "modified.xlsx";

            // Load the original workbook
            Workbook workbook = new Workbook(sourcePath);

            // Example modification: change a cell value
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Modified content");

            // Save the modified workbook to the new file (preserves format)
            workbook.Save(destinationPath, SaveFormat.Xlsx);

            // Preserve the original version metadata (document properties)
            // Create MetadataOptions for document properties
            MetadataOptions metadataOptions = new MetadataOptions(MetadataType.DocumentProperties);

            // Load metadata from the original workbook
            WorkbookMetadata metadata = new WorkbookMetadata(sourcePath, metadataOptions);

            // Save the metadata to the new workbook file
            metadata.Save(destinationPath);

            Console.WriteLine("Workbook saved with original metadata preserved.");
        }
    }
}
