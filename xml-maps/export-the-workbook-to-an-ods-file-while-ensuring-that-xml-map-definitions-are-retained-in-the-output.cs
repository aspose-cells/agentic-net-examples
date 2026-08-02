using System;
using Aspose.Cells;
using Aspose.Cells.Ods;

class ExportWorkbookWithXmlMapToOds
{
    static void Main()
    {
        // Load a workbook that already contains XML map definitions
        Workbook workbook = new Workbook("InputWithXmlMap.xlsx");

        // Verify that at least one XML map exists (optional)
        if (workbook.Worksheets.XmlMaps.Count == 0)
        {
            Console.WriteLine("Warning: No XML maps found in the workbook.");
        }

        // Create ODS save options
        OdsSaveOptions odsOptions = new OdsSaveOptions();
        odsOptions.GeneratorType = OdsGeneratorType.LibreOffice; // set generator (optional)

        // Save the workbook as ODS; XML map definitions are preserved automatically
        workbook.Save("OutputWithXmlMap.ods", odsOptions);

        Console.WriteLine("Workbook successfully saved as ODS with XML maps retained.");
    }
}