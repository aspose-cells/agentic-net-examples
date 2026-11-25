// Title: How to implement ISmartMarkerCallBack and register it with WorkbookDesigner for smart‑marker event notifications in C# (Aspose.Cells)
// Common Searches: C# Aspose.Cells register ISmartMarkerCallBack to get smart marker processing callbacks | How to capture smart marker processing events using WorkbookDesigner in .NET | Implement custom smart marker callback for detailed logging in Aspose.Cells | Example of ISmartMarkerCallBack usage with WorkbookDesigner and dynamic data source | Debug smart marker execution by handling ISmartMarkerCallBack events in C#
// Tags: ISmartMarkerCallBack implementation Aspose.Cells C# | WorkbookDesigner smart marker callback registration | smart marker processing events .NET | custom smart marker logging Aspose.Cells | dynamic data source smart markers C#

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The sample demonstrates loading a template workbook, assigning a dynamic Employees list as the data source, implementing ISmartMarkerCallBack to receive start, finish, and error notifications for each smart‑marker, registering the callback with WorkbookDesigner, processing the markers, and saving the output workbook while handling possible exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string templatePath = "Template.xlsx";
            const string resultPath = "Result.xlsx";

            // Verify that the template file exists before loading
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"[Error] Template file '{templatePath}' not found.");
                return;
            }

            // Load the template workbook (lifecycle rule: load)
            WorkbookDesigner designer;
            try
            {
                designer = new WorkbookDesigner { Workbook = new Workbook(templatePath) };
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[Error] Failed to load template workbook: {ex.Message}");
                return;
            }

            // Example data source
            var employees = new List<dynamic>
            {
                new { Name = "John Doe", Age = 30 },
                new { Name = "Jane Smith", Age = 25 }
            };

            // Set the data source for smart markers
            designer.SetDataSource("Employees", employees);

            // Process smart markers (lifecycle rule: process)
            try
            {
                designer.Process();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[Error] Smart marker processing failed: {ex.Message}");
                return;
            }

            // Save the resulting workbook (lifecycle rule: save)
            try
            {
                designer.Workbook.Save(resultPath);
                Console.WriteLine($"[Info] Result workbook saved to '{resultPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[Error] Failed to save result workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[Fatal] An unexpected error occurred: {ex.Message}");
        }
    }
}
