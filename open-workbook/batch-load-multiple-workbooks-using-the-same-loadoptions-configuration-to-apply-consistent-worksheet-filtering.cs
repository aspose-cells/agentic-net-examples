using System;
using System.IO;
using Aspose.Cells;

// Author: Aspose.Cells .NET example – batch load workbooks with a shared LoadFilter

// Custom LoadFilter to control which data is loaded per worksheet
public class CustomLoadFilter : LoadFilter
{
    // This method is called before each worksheet is loaded.
    // Adjust LoadDataFilterOptions based on worksheet name or other criteria.
    public override void StartSheet(Worksheet sheet)
    {
        // Example: load full data for sheets named "Data", otherwise only structure.
        if (sheet.Name.Equals("Data", StringComparison.OrdinalIgnoreCase))
        {
            LoadDataFilterOptions = LoadDataFilterOptions.All;
        }
        else
        {
            LoadDataFilterOptions = LoadDataFilterOptions.Structure;
        }
    }
}

class Program
{
    static void Main()
    {
        // Files to be processed
        string[] workbookFiles = new string[]
        {
            "Workbook1.xlsx",
            "Workbook2.xlsx",
            "Workbook3.xlsx"
        };

        // Create a single LoadOptions instance and assign the custom filter
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LoadFilter = new CustomLoadFilter();

        // Load each workbook using the same LoadOptions configuration
        foreach (string filePath in workbookFiles)
        {
            // Load workbook with the shared LoadOptions
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Example operation: output worksheet count
            Console.WriteLine($"{Path.GetFileName(filePath)} loaded with {workbook.Worksheets.Count} worksheets.");

            // Save the processed workbook (optional)
            string outputPath = $"Processed_{Path.GetFileName(filePath)}";
            workbook.Save(outputPath);
        }
    }
}