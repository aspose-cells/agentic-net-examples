// Title: C# wrapper for Aspose.Cells Workbook that logs constructor, Load and Save calls to a debug file for localization troubleshooting
// Common Searches: how to log Aspose.Cells workbook Save overload parameters in C# | trace Load and Save method calls of Aspose.Cells for localization debugging | create a thread‑safe wrapper around Aspose.Cells Workbook to write method logs | debug runtime localization issues in Aspose.Cells by logging workbook operations | C# example of writing Aspose.Cells method call trace to a text file
// Tags: Aspose.Cells workbook method call logging C# | thread‑safe file logging .NET | wrapper class for Aspose.Cells Workbook | debug localization runtime Aspose.Cells | log Save overload parameters Aspose.Cells | trace Load operation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The LoggingWorkbook class acts as a C# wrapper around Aspose.Cells Workbook, recording each constructor, static Load, and all Save overload invocations to a thread‑safe log file (method_calls.log) with timestamps and parameter details. It forwards the actual workbook operations to the underlying Workbook while exposing the Worksheets collection and the raw Workbook for advanced usage, enabling developers to troubleshoot localization behavior at runtime.
public class LoggingWorkbook
{
    private readonly Workbook _workbook;
    private static readonly string LogFilePath = "method_calls.log";
    private static readonly object LogLock = new object();

    private static void Log(string methodName, params string[] parameters)
    {
        lock (LogLock)
        {
            string entry = $"{DateTime.UtcNow:O} - {methodName}({string.Join(", ", parameters)}){Environment.NewLine}";
            File.AppendAllText(LogFilePath, entry);
        }
    }

    // Default constructor – creates a new empty workbook
    public LoggingWorkbook()
    {
        Log(nameof(LoggingWorkbook) + ".ctor", "default");
        _workbook = new Workbook();
    }

    // Private constructor used by the Load method
    private LoggingWorkbook(Workbook workbook)
    {
        Log(nameof(LoggingWorkbook) + ".ctor", "loaded");
        _workbook = workbook;
    }

    // Static Load method – logs the call and returns a wrapped workbook
    public static LoggingWorkbook Load(string fileName)
    {
        Log(nameof(Load), $"fileName=\"{fileName}\"");
        Workbook wb = new Workbook(fileName);
        return new LoggingWorkbook(wb);
    }

    // Save overloads – each logs its parameters before delegating to the underlying workbook
    public void Save(string fileName)
    {
        Log(nameof(Save), $"fileName=\"{fileName}\"");
        _workbook.Save(fileName);
    }

    public void Save(string fileName, SaveFormat format)
    {
        Log(nameof(Save), $"fileName=\"{fileName}\", format={format}");
        _workbook.Save(fileName, format);
    }

    public void Save(Stream stream, SaveFormat format)
    {
        Log(nameof(Save), $"stream, format={format}");
        _workbook.Save(stream, format);
    }

    // Expose the Worksheets collection for normal workbook operations
    public WorksheetCollection Worksheets => _workbook.Worksheets;

    // Expose the underlying Workbook in case advanced features are needed
    public Workbook UnderlyingWorkbook => _workbook;
}

// Example usage demonstrating creation, modification, saving, and loading with logging
class Program
{
    static void Main()
    {
        // Create a new workbook (log constructor)
        var wb = new LoggingWorkbook();

        // Write some data
        wb.Worksheets[0].Cells["A1"].PutValue("Localized Text");

        // Save the workbook (log Save)
        wb.Save("Sample.xlsx");

        // Load the workbook (log Load)
        var loadedWb = LoggingWorkbook.Load("Sample.xlsx");

        // Verify that data is accessible after load
        string value = loadedWb.Worksheets[0].Cells["A1"].StringValue;
        Console.WriteLine($"Loaded cell value: {value}");
    }
}
