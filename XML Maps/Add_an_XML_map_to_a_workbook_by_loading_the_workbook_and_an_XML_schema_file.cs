using System;
using Aspose.Cells;

public class Program
{
    public static void Main()
    {
        AddXmlMapExample.Run();
    }
}

public class AddXmlMapExample
{
    public static void Run()
    {
        string inputPath = "input.xlsx";
        string schemaPath = "schema.xsd";
        string outputPath = "output_with_xmlmap.xlsx";

        Workbook workbook = new Workbook(inputPath);
        int mapIndex = workbook.Worksheets.XmlMaps.Add(schemaPath);
        XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
        xmlMap.Name = "MyXmlMap";

        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}