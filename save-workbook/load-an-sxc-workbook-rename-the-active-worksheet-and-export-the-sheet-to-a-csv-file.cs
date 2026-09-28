// Title: Load an SXC workbook, rename the active worksheet, and export it to CSV using Aspose.Cells for .NET
// AI Prompts: Open a .sxc file with Aspose.Cells, change the name of the currently active worksheet, and save that sheet as a .csv file in C#. | Using Aspose.Cells for .NET, load an SXC workbook, assign a new name to the active sheet, and export the first worksheet to CSV format. | Show how to rename the active worksheet after loading an SXC workbook and then convert the workbook to a CSV file with Aspose.Cells.
// Common Searches: C# Aspose.Cells rename active sheet after loading SXC and export to CSV | Convert first worksheet of an SXC workbook to CSV after renaming it using Aspose.Cells .NET | How to change worksheet name in an SXC file before saving as CSV with Aspose.Cells
// Tags: Aspose.Cells load SXC file | active sheet rename operation Aspose.Cells | worksheet to CSV conversion Aspose.Cells | SaveFormat.Csv parameter usage Aspose.Cells | modify worksheet name prior to CSV export .NET

using System;
using Aspose.Cells;

// // Loads an SXC workbook, renames the active worksheet to "RenamedSheet", and saves the first worksheet as a CSV file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Path to the source SXC workbook
        string sourcePath = "input.sxc";

        // Path for the exported CSV file
        string csvPath = "output.csv";

        // Load the SXC workbook from file
        Workbook workbook = new Workbook(sourcePath);

        // Get the index of the active worksheet
        int activeIndex = workbook.Worksheets.ActiveSheetIndex;

        // Rename the active worksheet
        workbook.Worksheets[activeIndex].Name = "RenamedSheet";

        // Export the workbook (first sheet) to CSV format
        workbook.Save(csvPath, SaveFormat.Csv);
    }
}
