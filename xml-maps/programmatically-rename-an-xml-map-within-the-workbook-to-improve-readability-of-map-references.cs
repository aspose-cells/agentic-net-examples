// Title: Renaming an XML map in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# rename XML map in existing workbook | How to change XmlMap name in an .xlsx file with Aspose.Cells | C# check for duplicate XmlMap names before renaming using Aspose.Cells | Programmatic XML map rename example with Aspose.Cells for .NET | Update XML map reference name in Excel workbook via Aspose.Cells API
// Tags: XmlMap name update Aspose.Cells | Excel workbook XML map conflict detection | Aspose.Cells XmlMap collection handling | C# modify XML map identifier and save workbook | prevent duplicate XmlMap names Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook (or creates one with a sample XML schema), accesses its XmlMapCollection, checks whether the desired new name already exists, updates the first XmlMap's Name property if no conflict is found, and saves the workbook to a new file, demonstrating safe XML map renaming with Aspose.Cells for .NET.
class Program
{
    static void Main(string[] args)
    {
        // Paths for input and output workbooks
        string inputPath = "input.xlsx";
        string outputPath = "output_renamed.xmlmap.xlsx";

        // Load an existing workbook if it exists; otherwise create a sample workbook with an XML map
        Workbook workbook;
        if (File.Exists(inputPath))
        {
            workbook = new Workbook(inputPath);
        }
        else
        {
            workbook = new Workbook();
            Worksheet ws = workbook.Worksheets[0];
            ws.Name = "DataSheet";

            // Sample data
            ws.Cells["A1"].PutValue("Name");
            ws.Cells["B1"].PutValue("Value");
            ws.Cells["A2"].PutValue("Item1");
            ws.Cells["B2"].PutValue(100);
            ws.Cells["A3"].PutValue("Item2");
            ws.Cells["B3"].PutValue(200);

            // Sample XML schema to create an XmlMap
            string xmlSchema = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
  <xs:element name='Data'>
    <xs:complexType>
      <xs:sequence>
        <xs:element name='Item' maxOccurs='unbounded'>
          <xs:complexType>
            <xs:sequence>
              <xs:element name='Name' type='xs:string'/>
              <xs:element name='Value' type='xs:integer'/>
            </xs:sequence>
          </xs:complexType>
        </xs:element>
      </xs:sequence>
    </xs:complexType>
  </xs:element>
</xs:schema>";

            // Add the XML map to the workbook
            int mapIndex = workbook.Worksheets.XmlMaps.Add(xmlSchema);
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
            xmlMap.Name = "OriginalMapName";
        }

        // Access the collection of XML maps in the workbook
        XmlMapCollection xmlMaps = workbook.Worksheets.XmlMaps;

        if (xmlMaps.Count == 0)
        {
            Console.WriteLine("No XML maps are present in the workbook.");
            return;
        }

        // For this example, rename the first XML map in the collection
        XmlMap mapToRename = xmlMaps[0];
        Console.WriteLine($"Current XML Map Name: {mapToRename.Name}");

        string desiredNewName = "RenamedMap";

        // Verify that the new name does not clash with an existing map name
        bool nameExists = false;
        foreach (XmlMap existingMap in xmlMaps)
        {
            if (existingMap.Name.Equals(desiredNewName, StringComparison.OrdinalIgnoreCase))
            {
                nameExists = true;
                break;
            }
        }

        if (nameExists)
        {
            Console.WriteLine($"An XML map named '{desiredNewName}' already exists. Choose a different name.");
        }
        else
        {
            mapToRename.Name = desiredNewName;
            Console.WriteLine($"XML Map successfully renamed to: {mapToRename.Name}");
        }

        // Save the modified workbook
        workbook.Save(outputPath);
        Console.WriteLine($"Workbook saved to '{outputPath}'.");
    }
}
