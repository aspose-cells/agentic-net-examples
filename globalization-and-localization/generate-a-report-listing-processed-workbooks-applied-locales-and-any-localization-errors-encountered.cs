// Title: Generate a localization report for multiple Excel workbooks with applied CultureInfo and error logging using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# apply French locale to Excel workbooks and log errors | create a summary workbook of processed files with culture info using Aspose.Cells | C# generate Excel report of workbook processing results and localization exceptions | save workbook after setting Settings.CultureInfo with Aspose.Cells .NET | capture and record localization errors when updating Excel file culture in C#
// Tags: Aspose.Cells workbook cultureinfo assignment | C# Excel localization report generation | Aspose.Cells error handling for culture settings | Excel file processing summary with Aspose.Cells | Aspose.Cells save report workbook

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// The example loads each Excel file from a list, assigns a target CultureInfo (e.g., fr-FR) to the workbook via Settings.CultureInfo, saves the workbook, and captures any exception. It records the workbook name, applied locale, and error message, then creates a new workbook, writes a header (Workbook, Applied Locale, Error) and populates rows with the collected data, finally saving the report as LocalizationReport.xlsx.
class Program
{
    static void Main()
    {
        // List of workbook file paths to process
        var workbookFiles = new List<string>
        {
            "Book1.xlsx",
            "Book2.xlsx"
            // Add additional workbook paths as needed
        };

        // Desired locale to apply to each workbook
        var targetLocale = new CultureInfo("fr-FR");

        // Collection to hold report entries
        var reportEntries = new List<(string WorkbookName, string Locale, string ErrorMessage)>();

        foreach (var filePath in workbookFiles)
        {
            string error = string.Empty;

            try
            {
                // Load workbook (use provided load rule)
                var wb = new Workbook(filePath);

                // Apply locale (use provided create rule by setting property)
                wb.Settings.CultureInfo = targetLocale;

                // Save workbook after applying locale (use provided save rule)
                wb.Save(filePath);
            }
            catch (Exception ex)
            {
                // Capture any localization errors
                error = ex.Message;
            }

            // Record the result for the report
            reportEntries.Add((Path.GetFileName(filePath), targetLocale.Name, error));
        }

        // Create a new workbook for the report (use provided create rule)
        var reportWb = new Workbook();
        var sheet = reportWb.Worksheets[0];
        sheet.Name = "Localization Report";

        // Write header row
        sheet.Cells["A1"].PutValue("Workbook");
        sheet.Cells["B1"].PutValue("Applied Locale");
        sheet.Cells["C1"].PutValue("Error");

        // Populate report rows
        for (int i = 0; i < reportEntries.Count; i++)
        {
            int rowIndex = i + 2; // Data starts at row 2
            sheet.Cells[$"A{rowIndex}"].PutValue(reportEntries[i].WorkbookName);
            sheet.Cells[$"B{rowIndex}"].PutValue(reportEntries[i].Locale);
            sheet.Cells[$"C{rowIndex}"].PutValue(reportEntries[i].ErrorMessage);
        }

        // Save the report workbook (use provided save rule)
        reportWb.Save("LocalizationReport.xlsx");
    }
}
