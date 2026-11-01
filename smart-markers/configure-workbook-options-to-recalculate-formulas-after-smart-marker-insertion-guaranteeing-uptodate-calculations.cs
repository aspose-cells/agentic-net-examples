// Title: How to recalculate all formulas after inserting smart markers with Aspose.Cells in C#
// AI Prompts: Use WorkbookDesigner to populate a workbook via smart markers, then call CalculateFormula to update every formula before saving. | Process smart markers in an Aspose.Cells workbook and trigger automatic formula refresh prior to export in C#. | Force immediate evaluation of Excel formulas after WorkbookDesigner.Process() using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# recalculate formulas after WorkbookDesigner.Process | force formula calculation after smart marker data insertion Aspose.Cells | how to update Excel formulas after using smart markers with Aspose.Cells .NET
// Tags: smart marker data binding with WorkbookDesigner | force formula calculation after smart marker processing | Aspose.Cells recalculate formulas C# | Excel template update using smart markers | automatic formula refresh Aspose.Cells

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// The example loads an Excel template, creates a DataTable of employee data, processes smart markers with WorkbookDesigner, forces immediate formula calculation using CalculateFormula, and saves the result, ensuring all formulas reflect the inserted data.
class Program
{
    static void Main()
    {
        const string templatePath = "SmartMarkerTemplate.xlsx";
        const string resultPath = "Result.xlsx";

        // Verify that the template file exists to avoid FileNotFoundException
        if (!File.Exists(templatePath))
        {
            Console.WriteLine($"Error: Template file \"{templatePath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook that contains smart markers
            Workbook workbook = new Workbook(templatePath);

            // Prepare a data source for the smart markers
            DataTable dt = new DataTable("Employees");
            dt.Columns.Add("Name");
            dt.Columns.Add("Salary", typeof(double));
            dt.Rows.Add("John Doe", 50000);
            dt.Rows.Add("Jane Smith", 60000);

            // Process the smart markers with the data source
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(dt);
            designer.Process();

            // Force immediate calculation of all formulas
            workbook.CalculateFormula();

            // Save the updated workbook
            workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved successfully to \"{resultPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
