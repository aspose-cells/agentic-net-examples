// Title: Log IFormatProvider and ICustomFormatter method calls to a debug file for Aspose.Cells localization troubleshooting in C#
// Common Searches: how to trace GetFormat and Format calls of a custom IFormatProvider in .NET | debug localization provider in Aspose.Cells by writing method calls to a log file | C# example of logging ICustomFormatter usage while formatting Excel cell values | record custom format provider activity to a file when generating an Excel workbook with Aspose.Cells | troubleshoot Aspose.Cells localization by capturing format provider method invocations
// Tags: custom IFormatProvider logging for Aspose.Cells | ICustomFormatter method tracing in C# | Aspose.Cells localization debug log | write format provider calls to file .NET | debugging localization with custom formatter | Excel workbook formatting using custom provider

using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsLocalizationDebug
{
    // Custom format provider that logs each method call for troubleshooting.
    // The sample defines a DebugLocalizationProvider that implements IFormatProvider and ICustomFormatter, logging each GetFormat and Format invocation to a specified file. The provider is passed to string.Format to format a date and a number, which are then written to cells of an Aspose.Cells workbook. Saving the workbook produces an Excel file while the debug log captures the localization provider's runtime behavior for troubleshooting.
    public class DebugLocalizationProvider : IFormatProvider, ICustomFormatter
    {
        private readonly string _logFilePath;

        public DebugLocalizationProvider(string logFilePath)
        {
            _logFilePath = logFilePath;
            // Ensure the log file exists.
            File.WriteAllText(_logFilePath, $"DebugLocalizationProvider initialized at {DateTime.Now}{Environment.NewLine}");
        }

        // IFormatProvider implementation.
        public object GetFormat(Type formatType)
        {
            Log($"GetFormat called with formatType = {formatType}");
            // Return this instance if the requested format is ICustomFormatter.
            if (formatType == typeof(ICustomFormatter))
                return this;
            // Otherwise defer to the default provider.
            return CultureInfo.CurrentCulture.GetFormat(formatType);
        }

        // ICustomFormatter implementation.
        public string Format(string format, object arg, IFormatProvider formatProvider)
        {
            Log($"Format called with format = \"{format}\", arg = {arg}, formatProvider = {formatProvider?.GetType().Name ?? "null"}");
            // Use the default formatting behavior.
            if (arg is IFormattable formattable)
                return formattable.ToString(format, CultureInfo.CurrentCulture);
            return arg?.ToString() ?? string.Empty;
        }

        // Helper method to write log entries.
        private void Log(string message)
        {
            string entry = $"{DateTime.Now:O} - {message}{Environment.NewLine}";
            File.AppendAllText(_logFilePath, entry);
            Debug.WriteLine(entry);
        }
    }

    class Program
    {
        static void Main()
        {
            // Path for the debug log.
            string logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "localization_debug.log");

            // Instantiate the custom provider.
            var debugProvider = new DebugLocalizationProvider(logPath);

            // Create a new workbook using Aspose.Cells lifecycle rule.
            var workbook = new Workbook(); // create

            // Access the first worksheet.
            var sheet = workbook.Worksheets[0];

            // Add some data that will be formatted using the custom provider.
            // Example: a date and a number.
            DateTime now = DateTime.Now;
            double number = 12345.6789;

            // Use string.Format with the custom provider to generate formatted strings.
            string dateString = string.Format(debugProvider, "{0:D}", now);
            string numberString = string.Format(debugProvider, "{0:N2}", number);

            // Write the formatted strings into cells.
            sheet.Cells["A1"].PutValue(dateString);
            sheet.Cells["A2"].PutValue(numberString);

            // Save the workbook (using the provided save rule).
            string outputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "LocalizedOutput.xlsx");
            workbook.Save(outputPath); // save

            // Indicate completion.
            Console.WriteLine($"Workbook saved to {outputPath}");
            Console.WriteLine($"Debug log written to {logPath}");
        }
    }
}
