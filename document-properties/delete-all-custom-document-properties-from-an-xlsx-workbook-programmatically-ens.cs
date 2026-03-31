using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Properties;

class Program
{
    static void Main(string[] args)
    {
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: <inputPath> <outputPath>");
            return;
        }

        DeleteCustomProperties.Run(args[0], args[1]);
    }
}

class DeleteCustomProperties
{
    public static void Run(string inputPath, string outputPath)
    {
        Workbook workbook = new Workbook(inputPath);
        CustomDocumentPropertyCollection customProps = workbook.CustomDocumentProperties;

        List<string> namesToRemove = new List<string>();
        foreach (DocumentProperty prop in customProps)
        {
            namesToRemove.Add(prop.Name);
        }

        foreach (string name in namesToRemove)
        {
            customProps.Remove(name);
        }

        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}