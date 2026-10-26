// Title: Rename an XML map in an Excel workbook with Aspose.Cells for .NET (C#)
// Common Searches: C# code to change the name of an existing XML map in an Excel file using Aspose.Cells | Aspose.Cells rename XML map programmatically in a .xlsx workbook | How to update the XML map name in a workbook with the Aspose.Cells .NET API | Set a readable name for an XML map in Excel using Aspose.Cells C# | Modify XML map reference after loading a workbook with Aspose.Cells
// Tags: aspose.cells xml map rename | c# set xml map name in workbook | aspose.cells modify xml map reference | excel .xlsx xml map name update using aspose | programmatic xml map renaming with aspose.cells

using Aspose.Cells;

// The example loads an Excel workbook, checks for XML maps, accesses the first map, assigns a more readable value to its Name property, and saves the workbook with the updated XML map name.
class Program
{
    static void Main()
    {
        // Load the workbook (load rule)
        var workbook = new Workbook("input.xlsx");

        // Check that the workbook contains at least one XML map
        if (workbook.Worksheets.XmlMaps.Count > 0)
        {
            // Access the XML map you want to rename (by index or by existing name)
            var xmlMap = workbook.Worksheets.XmlMaps[0];

            // Rename the XML map to improve readability
            xmlMap.Name = "ReadableMapName";
        }

        // Save the workbook with the updated XML map name (save rule)
        workbook.Save("output.xlsx");
    }
}
