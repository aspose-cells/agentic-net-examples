using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class AssignMacroFromTsv
{
    static void Main()
    {
        // Path to the TSV file.
        string tsvPath = "macros.tsv";

        // If the TSV file does not exist, create a sample one.
        if (!File.Exists(tsvPath))
        {
            File.WriteAllText(tsvPath,
                "Rectangle\t1\t1\t100\t100\tDoWork()\n" +
                "Oval\t3\t2\t80\t120\tDoOtherWork()");
        }

        // Create a new workbook and enable macros.
        Workbook workbook = new Workbook();
        workbook.Settings.EnableMacros = true;
        Worksheet sheet = workbook.Worksheets[0];

        // Read TSV lines.
        foreach (string line in File.ReadAllLines(tsvPath))
        {
            if (string.IsNullOrWhiteSpace(line))
                continue;

            string[] parts = line.Split('\t');
            if (parts.Length < 6)
                continue;

            string shapeType = parts[0].Trim();
            int row = int.Parse(parts[1]);
            int column = int.Parse(parts[2]);
            int width = int.Parse(parts[3]);
            int height = int.Parse(parts[4]);
            string macroName = parts[5].Trim();

            Shape shape = null;

            // Add shape based on the requested type.
            if (shapeType.Equals("Rectangle", StringComparison.OrdinalIgnoreCase))
            {
                // Parameters: upperLeftRow, upperLeftColumn, upperLeftRowOffset, upperLeftColumnOffset, height, width
                shape = sheet.Shapes.AddRectangle(row, column, 0, 0, height, width);
            }
            else if (shapeType.Equals("Oval", StringComparison.OrdinalIgnoreCase))
            {
                shape = sheet.Shapes.AddOval(row, column, 0, 0, height, width);
            }
            // Additional shape types can be added here.

            if (shape != null)
            {
                shape.MacroName = macroName;
                Console.WriteLine($"Assigned macro '{macroName}' to {shapeType} at ({row},{column}).");
            }
        }

        // Save the workbook as a macro‑enabled file.
        workbook.Save("WorkbookWithMacros.xlsm", SaveFormat.Xlsm);
        Console.WriteLine("Workbook saved as WorkbookWithMacros.xlsm");
    }
}