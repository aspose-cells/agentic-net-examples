// Title: Detect unresolved Smart Markers after processing a workbook with Aspose.Cells for .NET (C#)
// Description: Loads a template workbook, binds a DataTable to WorkbookDesigner, runs Process(), calls GetSmartMarkers() to identify any remaining markers, logs the results, and saves the final file.
// Keywords: Aspose.Cells smart marker validation | C# GetSmartMarkers example | WorkbookDesigner unresolved markers | smart marker resolution .NET | check smart markers after processing
// Common Searches: Aspose.Cells GetSmartMarkers C# example | how to verify smart markers are resolved | detect unresolved smart markers in Excel using Aspose | smart marker validation after WorkbookDesigner.Process | C# check for remaining smart markers
// Developer Intent: Confirm that no smart marker tags are left unresolved after WorkbookDesigner.Process runs.
// Use Cases: Automated quality gate: fail a CI build if GetSmartMarkers returns any items. | Debugging data mismatches by logging each unresolved marker before saving the workbook. | Batch processing of multiple templates where each must be fully resolved before distribution.
// AI Prompts: Generate C# code that loads an Excel template, binds several DataTables, processes smart markers with WorkbookDesigner, and throws an exception when GetSmartMarkers reports any unresolved tags. | Show how to iterate over GetSmartMarkers results, write each marker to a log file, and continue processing only if the list is empty.

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// Loads a template workbook, binds a DataTable to WorkbookDesigner, runs Process(), calls GetSmartMarkers() to identify any remaining markers, logs the results, and saves the final file.
public class SmartMarkerValidation
{
    public static void Main()
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }

    public static void Run()
    {
        const string templatePath = "template.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the template file exists to avoid FileNotFoundException
        if (!File.Exists(templatePath))
        {
            Console.WriteLine($"Template file not found: {templatePath}");
            return;
        }

        try
        {
            // Load the template workbook that contains smart markers
            Workbook templateWorkbook = new Workbook(templatePath);

            // Initialize the WorkbookDesigner with the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = templateWorkbook
            };

            // Create a sample data source (DataTable) for demonstration
            DataTable dataTable = new DataTable("Employees");
            dataTable.Columns.Add("Name", typeof(string));
            dataTable.Columns.Add("Age", typeof(int));
            dataTable.Rows.Add("John Doe", 30);
            dataTable.Rows.Add("Jane Smith", 28);

            // Bind the data source to the designer
            designer.SetDataSource(dataTable);

            // Process the smart markers in the workbook
            designer.Process();

            // Retrieve any smart markers that remain unresolved after processing
            string[] unresolvedMarkers = designer.GetSmartMarkers();

            // Report the validation result
            if (unresolvedMarkers.Length == 0)
            {
                Console.WriteLine("All smart markers have been resolved.");
            }
            else
            {
                Console.WriteLine("Unresolved smart markers detected:");
                foreach (string marker in unresolvedMarkers)
                {
                    Console.WriteLine(marker);
                }
            }

            // Save the processed workbook
            designer.Workbook.Save(outputPath);
            Console.WriteLine($"Processed workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during processing: {ex.Message}");
        }
    }
}
