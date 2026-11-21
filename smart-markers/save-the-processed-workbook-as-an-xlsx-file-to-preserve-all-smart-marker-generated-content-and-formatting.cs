// Title: Saving a workbook with processed Smart Markers to XLSX using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to export a WorkbookDesigner processed workbook to XLSX | preserve smart marker formatting when saving workbook as XLSX in .NET | Aspose.Cells save workbook after smart marker processing with custom output folder | example code for saving processed smart markers to an XLSX file using Aspose.Cells | Aspose.Cells WorkbookDesigner output file path and format selection in C#
// Tags: Aspose.Cells WorkbookDesigner save as XLSX | C# smart marker processing output file | retain smart marker styling Aspose.Cells | create result directory before workbook save C# | Aspose.Cells save format options .NET | smart markers to Excel file using Aspose.Cells

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// The example loads a template workbook containing Smart Markers, builds a DataTable as the data source, processes the markers with WorkbookDesigner, ensures the target directory exists, and then saves the resulting workbook as an XLSX file, preserving all generated content and formatting.
class Program
{
    static void Main()
    {
        try
        {
            const string templatePath = "Template.xlsx";
            const string resultPath = "Result.xlsx";

            // Verify that the template file exists before loading.
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Error: Template file \"{templatePath}\" not found.");
                return;
            }

            // Load the workbook that contains smart markers.
            Workbook workbook = new Workbook(templatePath);

            // Create a data source for the smart markers.
            DataTable data = new DataTable();
            data.Columns.Add("Name", typeof(string));
            data.Columns.Add("Age", typeof(int));
            data.Rows.Add("John", 30);
            data.Rows.Add("Jane", 28);

            // Process the smart markers using WorkbookDesigner.
            try
            {
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource(data);
                designer.Process();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Smart marker processing failed: {ex.Message}");
                return;
            }

            // Ensure the output directory exists.
            string resultDir = Path.GetDirectoryName(Path.GetFullPath(resultPath));
            if (!string.IsNullOrEmpty(resultDir) && !Directory.Exists(resultDir))
            {
                Directory.CreateDirectory(resultDir);
            }

            // Save the processed workbook.
            workbook.Save(resultPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to \"{resultPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
