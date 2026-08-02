// Title: Log Aspose.Cells Load Warnings to a File Using C# IWarningCallback
// Description: Demonstrates how to capture load‑time warnings about unsupported Excel features with Aspose.Cells. A custom IWarningCallback writes each warning's type and description to a specified log file via LoadOptions, enabling post‑load analysis and compatibility tracking in .NET applications.
// Keywords: Aspose.Cells | C# | IWarningCallback | LoadOptions | log warnings | unsupported features | Excel workbook loading | file logging | warning callback example | .NET | Aspose.Cells load warning
// Common Searches: Aspose.Cells log load warnings C# | IWarningCallback example .NET | capture unsupported feature warnings Aspose | write Aspose.Cells warnings to text file | LoadOptions WarningCallback usage
// Developer Intent: Capture every warning generated while opening an Excel workbook with Aspose.Cells and persist the messages to an application log for later review.
// Use Cases: Create an audit trail of compatibility issues when processing user‑uploaded spreadsheets. | Integrate warning logging into automated build or CI pipelines to flag unsupported Excel elements. | Provide detailed diagnostics for support teams investigating workbook import failures.
// AI Prompts: Generate a reusable IWarningCallback that logs to any ILogger implementation instead of a hard‑coded file path. | Show how to filter warnings by WarningInfo.Type == WarningType.UnsupportedFeature before writing to the log. | Write a script that parses the generated log file and outputs a summary count for each warning type.

using System;
using System.IO;
using Aspose.Cells;

// Custom warning callback that writes each warning to a log file
// Demonstrates how to capture load‑time warnings about unsupported Excel features with Aspose.Cells. A custom IWarningCallback writes each warning's type and description to a specified log file via LoadOptions, enabling post‑load analysis and compatibility tracking in .NET applications.
public class LoadWarningLogger : IWarningCallback
{
    private readonly string _logFilePath;

    public LoadWarningLogger(string logFilePath)
    {
        _logFilePath = logFilePath;
    }

    // This method is called by Aspose.Cells whenever a warning occurs during loading
    public void Warning(WarningInfo warningInfo)
    {
        string logEntry = $"Warning Type: {warningInfo.Type}, Description: {warningInfo.Description}";
        File.AppendAllText(_logFilePath, logEntry + Environment.NewLine);
    }
}

public class Program
{
    public static void Main()
    {
        // Path to the workbook that may contain unsupported features
        string inputFile = "input.xlsx";

        // Path to the application log file where warnings will be recorded
        string logFile = "load_warnings.log";

        // Create LoadOptions and assign the custom warning callback
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.WarningCallback = new LoadWarningLogger(logFile);
        loadOptions.IgnoreUselessShapes = true; // optional: ignore shapes that are not needed

        // Load the workbook with the specified options (warnings will be captured)
        Workbook workbook = new Workbook(inputFile, loadOptions);

        // Perform any additional processing if required
        // ...

        // Save the workbook (optional, demonstrates normal workflow)
        workbook.Save("output.xlsx");
    }
}
