using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace RenameTempNamedRanges
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Get the collection of defined names in the workbook
            NameCollection names = workbook.Worksheets.Names;

            // First, collect the names that start with "Temp"
            List<string> tempNames = new List<string>();
            foreach (Name name in names)
            {
                if (name.Text != null && name.Text.StartsWith("Temp", StringComparison.Ordinal))
                {
                    tempNames.Add(name.Text);
                }
            }

            // Rename each collected name by prefixing "Archive_"
            foreach (string oldName in tempNames)
            {
                Name name = names[oldName];
                name.Text = "Archive_" + oldName;
            }

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("output.xlsx");
        }
    }
}