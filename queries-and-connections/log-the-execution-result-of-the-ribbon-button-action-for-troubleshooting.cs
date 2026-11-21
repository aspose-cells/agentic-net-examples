// Title: Log execution of a custom Excel ribbon button that adds a worksheet using Aspose.Cells for .NET (C#)
// Common Searches: c# Aspose.Cells log ribbon button execution to text file | how to write timestamped log when adding worksheet via ribbon in Aspose.Cells | Aspose.Cells .NET check file existence before loading workbook in ribbon handler | catch exceptions and log errors for Excel ribbon custom action using Aspose.Cells | save modified workbook and create log entry after ribbon button click in C#
// Tags: ribbon button logging Aspose.Cells C# | add worksheet and save workbook Aspose.Cells | timestamped text file logging .NET | file existence validation before loading Excel workbook | exception handling with Aspose.Cells operations

using System;
using System.IO;
using Aspose.Cells;

// The example shows a RibbonHandler class that validates the source Excel file, loads it with Aspose.Cells, adds a new worksheet, saves the workbook, and writes a timestamped success or error message to a text log, including handling missing files and unexpected exceptions.
public class RibbonHandler
{
    // This method should be bound to the ribbon button click event.
    public void OnRibbonButtonClick()
    {
        // Paths for the source workbook, the result workbook, and the log file.
        string sourcePath = "input.xlsx";
        string resultPath = "output.xlsx";
        string logPath = "ribbon_log.txt";

        try
        {
            // Verify that the source file exists to avoid FileNotFoundException.
            if (!File.Exists(sourcePath))
            {
                string notFoundMsg = $"{DateTime.Now}: Source file '{sourcePath}' not found.{Environment.NewLine}";
                File.AppendAllText(logPath, notFoundMsg);
                return;
            }

            // Load the existing workbook.
            Workbook workbook = new Workbook(sourcePath);

            // Example operation: add a new worksheet named "LogSheet".
            int newSheetIndex = workbook.Worksheets.Add();
            workbook.Worksheets[newSheetIndex].Name = "LogSheet";

            // Save the modified workbook.
            workbook.Save(resultPath);

            // Log successful execution with timestamp.
            string successMessage = $"{DateTime.Now}: Ribbon button executed successfully. Added worksheet at index {newSheetIndex}.{Environment.NewLine}";
            File.AppendAllText(logPath, successMessage);
        }
        catch (Exception ex)
        {
            // Log any errors that occur during execution.
            string errorMessage = $"{DateTime.Now}: Ribbon button execution failed. Error: {ex.Message}{Environment.NewLine}";
            File.AppendAllText(logPath, errorMessage);
        }
    }

    // Entry point for the console application.
    public static void Main(string[] args)
    {
        try
        {
            RibbonHandler handler = new RibbonHandler();
            handler.OnRibbonButtonClick();
        }
        catch (Exception ex)
        {
            // Fallback logging if unexpected exception occurs.
            string fallbackLog = $"Unexpected error: {ex.Message}{Environment.NewLine}";
            File.AppendAllText("ribbon_log.txt", fallbackLog);
        }
    }
}
