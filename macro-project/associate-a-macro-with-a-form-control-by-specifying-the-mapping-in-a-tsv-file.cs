using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsMacroMapping
{
    class Program
    {
        static void Main()
        {
            string templatePath = "template.xlsm";
            Workbook workbook;

            if (File.Exists(templatePath))
            {
                workbook = new Workbook(templatePath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Sheet1";

                // Add a placeholder shape so the program can run even without a template
                Shape placeholder = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 1, 0, 0, 100, 30);
                placeholder.Name = "PlaceholderShape";
            }

            string tsvPath = "macro_mappings.tsv";

            if (!File.Exists(tsvPath))
            {
                Console.WriteLine($"TSV file \"{tsvPath}\" not found.");
                return;
            }

            foreach (string line in File.ReadAllLines(tsvPath))
            {
                if (string.IsNullOrWhiteSpace(line))
                    continue;

                string[] parts = line.Split('\t');
                if (parts.Length != 2)
                    continue;

                string shapeName = parts[0].Trim();
                string macroName = parts[1].Trim();

                Worksheet sheet = workbook.Worksheets[0];
                Shape shape = sheet.Shapes[shapeName];

                if (shape == null)
                {
                    Console.WriteLine($"Shape \"{shapeName}\" not found.");
                    continue;
                }

                shape.MacroName = macroName;
                Console.WriteLine($"Assigned macro \"{macroName}\" to shape \"{shapeName}\".");
            }

            workbook.Save("output.xlsm", SaveFormat.Xlsm);
        }
    }
}