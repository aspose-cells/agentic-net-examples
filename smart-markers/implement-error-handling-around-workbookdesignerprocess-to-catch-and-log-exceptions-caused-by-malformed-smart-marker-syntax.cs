// Title: Add try‑catch around WorkbookDesigner.Process to log malformed smart marker syntax errors in Aspose.Cells for .NET
// AI Prompts: Wrap the WorkbookDesigner.Process call in a try‑catch block that writes the exception message to the console or a log file. | Update SmartMarkerProcessor so that when designer.Process throws an exception, the method logs the error and returns early. | Preserve existing load and save error handling while adding detailed logging for any exception raised during smart marker processing.
// Common Searches: how to handle exceptions thrown by WorkbookDesigner.Process in Aspose.Cells C# | logging errors for invalid smart marker syntax using Aspose.Cells | example of try catch around smart marker processing in .NET | Aspose.Cells smart markers malformed syntax exception handling | C# code to catch and log smart marker processing failures
// Tags: WorkbookDesigner process exception handling | smart marker syntax error logging | Aspose.Cells smart marker error handling | C# try-catch for smart marker processing | Excel workbook smart marker processing failure

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, creates a WorkbookDesigner, and processes smart markers inside a try‑catch block that logs any exception caused by malformed smart marker syntax before saving the workbook, with additional error handling for loading and saving.
public class SmartMarkerProcessor
{
    /// <param name="inputPath">Path to the source Excel file.</param>
    /// <param name="outputPath">Path where the processed file will be saved.</param>
    public static void ProcessWorkbook(string inputPath, string outputPath)
    {
        // Verify input file exists
        if (!File.Exists(inputPath))
        {
            Console.Error.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook (lifecycle rule: load)
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Initialize the WorkbookDesigner (lifecycle rule: create)
        WorkbookDesigner designer = new WorkbookDesigner
        {
            Workbook = workbook
        };

        // Optional: set data source for smart markers here
        // designer.SetDataSource("Data", yourDataObject);

        try
        {
            // Process smart markers (target operation)
            designer.Process();
        }
        catch (Exception ex)
        {
            // Log the exception caused by malformed smart marker syntax
            Console.Error.WriteLine($"Smart marker processing error: {ex.Message}");
            return;
        }

        try
        {
            // Save the processed workbook (lifecycle rule: save)
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: SmartMarkerProcessor <inputPath> <outputPath>");
            return;
        }

        string inputPath = args[0];
        string outputPath = args[1];

        SmartMarkerProcessor.ProcessWorkbook(inputPath, outputPath);
    }
}
