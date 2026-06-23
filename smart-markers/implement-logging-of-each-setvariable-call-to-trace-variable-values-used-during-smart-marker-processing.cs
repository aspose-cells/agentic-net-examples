using System;
using System.IO;
using Aspose.Cells;

class SmartMarkerVariableLogger
{
    static void Main()
    {
        const string templatePath = "Template.xlsx";
        const string outputPath = "Output.xlsx";

        // Verify that the template file exists before loading
        if (!File.Exists(templatePath))
        {
            Console.WriteLine($"Template file not found: {templatePath}");
            return;
        }

        try
        {
            // Load the template workbook that contains smart markers
            Workbook workbook = new Workbook(templatePath);

            // Initialize the WorkbookDesigner with the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Prepare data object for smart markers
            var reportData = new ReportData
            {
                ReportDate = DateTime.Now,
                Title = "Sales Summary",
                Total = 12345.67
            };

            // Log each variable value
            LogSetVariable(nameof(reportData.ReportDate), reportData.ReportDate);
            LogSetVariable(nameof(reportData.Title), reportData.Title);
            LogSetVariable(nameof(reportData.Total), reportData.Total);

            // Set the data source for smart markers (named source)
            designer.SetDataSource("ReportData", reportData);

            // Process the smart markers using the defined variables
            designer.Process();

            // Save the processed workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Helper method to log each variable assignment
    static void LogSetVariable(string name, object value)
    {
        Console.WriteLine($"SetVariable called - Name: {name}, Value: {(value ?? "null")}");
    }

    // Simple data class used as the data source for smart markers
    private class ReportData
    {
        public DateTime ReportDate { get; set; }
        public string Title { get; set; } = null!;
        public double Total { get; set; }
    }
}