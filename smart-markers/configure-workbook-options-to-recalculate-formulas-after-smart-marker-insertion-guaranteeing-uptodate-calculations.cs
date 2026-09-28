// Title: Configure Aspose.Cells to recalculate formulas after processing smart markers in a C# workbook
// AI Prompts: Write C# code that loads an Excel template containing smart markers, binds a DataTable to a WorkbookDesigner, sets Designer.CalculateFormula to true, processes the markers, enables Workbook.Settings.FormulaSettings.CalculateOnSave, and saves the updated file. | Show how to programmatically force all formulas to refresh after inserting smart markers using Aspose.Cells WorkbookDesigner and FormulaSettings in a .NET application.
// Common Searches: Aspose.Cells how to refresh formulas after smart marker processing in C# | C# set calculate on save for workbook with smart markers | Enable automatic formula calculation when using WorkbookDesigner Aspose.Cells | Smart marker data binding recalculate formulas Aspose.Cells .NET example
// Tags: WorkbookDesigner calculate formulas | FormulaSettings calculate on save | smart marker data binding C# | Aspose.Cells recalculate formulas after processing | Excel smart markers automatic calculation

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example demonstrates loading an Excel template with smart markers, creating a DataTable of employee data, binding it to a WorkbookDesigner, enabling automatic formula calculation, processing the markers, configuring the workbook to recalculate formulas on save, and saving the result to a new file.
    public class SmartMarkerRecalculateDemo
    {
        public static void Run()
        {
            try
            {
                string templatePath = "SmartMarkerTemplate.xlsx";

                // Verify template file exists
                if (!File.Exists(templatePath))
                {
                    Console.WriteLine($"Template file not found: {templatePath}");
                    return;
                }

                // Load a workbook that contains smart markers
                Workbook workbook = new Workbook(templatePath);

                // Prepare a simple data source (DataTable) for demonstration
                DataTable dt = new DataTable("Employees");
                dt.Columns.Add("Name");
                dt.Columns.Add("Salary", typeof(double));
                dt.Rows.Add("John Doe", 5000);
                dt.Rows.Add("Jane Smith", 6200);

                // Initialize the WorkbookDesigner with the loaded workbook
                WorkbookDesigner designer = new WorkbookDesigner(workbook);

                // Bind the data source to the designer
                designer.SetDataSource(dt);
                designer.CalculateFormula = true; // Enable automatic formula calculation

                // Process all smart markers in the workbook
                designer.Process();

                // Ensure that formulas are also recalculated when the workbook is saved
                workbook.Settings.FormulaSettings.CalculateOnSave = true;

                // Save the resulting workbook
                string outputPath = "SmartMarkerRecalculated.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
