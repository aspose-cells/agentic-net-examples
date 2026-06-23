using System;
using System.IO;
using Aspose.Cells;

class CustomLoadFilter : LoadFilter
{
    // Load all data for each worksheet (including defined names)
    public override void StartSheet(Worksheet sheet)
    {
        this.LoadDataFilterOptions = LoadDataFilterOptions.All;
    }
}

class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Set load options with the custom filter
            var loadOptions = new LoadOptions
            {
                LoadFilter = new CustomLoadFilter()
            };

            // Load workbook using the specified options
            var workbook = new Workbook(inputPath, loadOptions);

            // Remove defined names that start with "temp_"
            NameCollection names = workbook.Worksheets.Names;
            for (int i = names.Count - 1; i >= 0; i--)
            {
                Name definedName = names[i];
                // The name string is accessed via the Text property
                if (definedName.Text.StartsWith("temp_", StringComparison.OrdinalIgnoreCase))
                {
                    names.RemoveAt(i);
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}