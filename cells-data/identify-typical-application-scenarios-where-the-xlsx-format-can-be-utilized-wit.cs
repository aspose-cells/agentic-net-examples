using System;
using Aspose.Cells;

namespace AsposeCellsXlsxScenarios
{
    class Program
    {
        static void Main(string[] args)
        {
            // Scenario 1: Generating a financial report
            // Create a new workbook (default format is Xlsx)
            Workbook reportWorkbook = new Workbook();

            // Add data to the first worksheet
            Worksheet reportSheet = reportWorkbook.Worksheets[0];
            reportSheet.Name = "FinancialReport";
            reportSheet.Cells["A1"].PutValue("Month");
            reportSheet.Cells["B1"].PutValue("Revenue");
            reportSheet.Cells["A2"].PutValue("January");
            reportSheet.Cells["B2"].PutValue(12500);
            reportSheet.Cells["A3"].PutValue("February");
            reportSheet.Cells["B3"].PutValue(15800);
            // Save as XLSX
            reportWorkbook.Save("FinancialReport.xlsx", SaveFormat.Xlsx);

            // Scenario 2: Exporting data from a database to Excel for user download
            // (In a real application, data would be read from a DB; here we mock it)
            Workbook exportWorkbook = new Workbook();
            Worksheet exportSheet = exportWorkbook.Worksheets[0];
            exportSheet.Name = "ExportData";
            exportSheet.Cells["A1"].PutValue("ID");
            exportSheet.Cells["B1"].PutValue("Name");
            exportSheet.Cells["A2"].PutValue(1);
            exportSheet.Cells["B2"].PutValue("Alice");
            exportSheet.Cells["A3"].PutValue(2);
            exportSheet.Cells["B3"].PutValue("Bob");
            exportWorkbook.Save("ExportedData.xlsx", SaveFormat.Xlsx);

            // Scenario 3: Creating a template file that end‑users will fill in
            Workbook templateWorkbook = new Workbook();
            Worksheet templateSheet = templateWorkbook.Worksheets[0];
            templateSheet.Name = "InputTemplate";
            templateSheet.Cells["A1"].PutValue("Enter your name");
            templateSheet.Cells["B1"].PutValue("Enter your email");
            // Mark the workbook as a template (Xltx) by saving with the appropriate format
            templateWorkbook.Save("InputTemplate.xltx", SaveFormat.Xltx);

            // Scenario 4: Performing data analysis and saving results
            Workbook analysisWorkbook = new Workbook();
            Worksheet analysisSheet = analysisWorkbook.Worksheets[0];
            analysisSheet.Name = "Analysis";
            // Sample data
            analysisSheet.Cells["A1"].PutValue("Value");
            analysisSheet.Cells["A2"].PutValue(10);
            analysisSheet.Cells["A3"].PutValue(20);
            analysisSheet.Cells["A4"].PutValue(30);
            // Simple formula
            analysisSheet.Cells["B1"].PutValue("Sum");
            analysisSheet.Cells["B2"].Formula = "SUM(A2:A4)";
            // Calculate formulas
            analysisWorkbook.CalculateFormula();
            // Save results
            analysisWorkbook.Save("AnalysisResults.xlsx", SaveFormat.Xlsx);

            // Scenario 5: Interchanging data with other Office applications (e.g., Word mail merge)
            Workbook mailMergeWorkbook = new Workbook();
            Worksheet mailMergeSheet = mailMergeWorkbook.Worksheets[0];
            mailMergeSheet.Name = "MailMergeData";
            mailMergeSheet.Cells["A1"].PutValue("FirstName");
            mailMergeSheet.Cells["B1"].PutValue("LastName");
            mailMergeSheet.Cells["A2"].PutValue("John");
            mailMergeSheet.Cells["B2"].PutValue("Doe");
            mailMergeSheet.Cells["A3"].PutValue("Jane");
            mailMergeSheet.Cells["B3"].PutValue("Smith");
            mailMergeWorkbook.Save("MailMergeData.xlsx", SaveFormat.Xlsx);

            Console.WriteLine("All example XLSX files have been created.");
        }
    }
}