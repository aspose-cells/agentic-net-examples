// Title: Add console logging for every WorkbookDesigner.SetVariable invocation while processing smart markers with Aspose.Cells in C#
// AI Prompts: Create a C# wrapper that logs the variable name and value to the console before calling WorkbookDesigner.SetDataSource during smart marker processing. | Show how to instrument the Aspose.Cells smart marker workflow so that each SetVariable call writes its parameters to a log file or console. | Demonstrate adding exception handling around SetVariable while preserving detailed logging of the variable data.
// Common Searches: how can I log the values passed to SetVariable in Aspose.Cells smart markers using C# | trace variable assignments during smart marker processing with Aspose.Cells WorkbookDesigner | debugging Aspose.Cells smart markers by printing SetVariable parameters to console | C# example of logging each SetVariable call before processing smart markers | Aspose.Cells SetVariable fallback to SetDataSource with logging
// Tags: Aspose.Cells log SetVariable calls | WorkbookDesigner variable tracing C# | smart marker debugging Aspose.Cells | setvariable console logging Aspose.Cells | smart marker variable logging C#

using System;
using System.IO;
using Aspose.Cells;

// The example wraps WorkbookDesigner.SetVariable in a method that writes the variable name and value to the console, then forwards the call to SetDataSource. This provides traceable logging for each variable used in Aspose.Cells smart marker processing, while preserving error handling and the ability to load templates, process markers, and save the workbook.
public class SmartMarkerProcessor
{
    private Workbook _workbook;
    private WorkbookDesigner _designer;

    // Constructor creates a new empty workbook and its designer
    public SmartMarkerProcessor()
    {
        try
        {
            _workbook = CreateWorkbook();
            _designer = new WorkbookDesigner(_workbook);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error initializing processor: {ex.Message}");
            throw;
        }
    }

    // Loads a template workbook if the file exists
    public void LoadTemplate(string templatePath)
    {
        try
        {
            if (!File.Exists(templatePath))
                throw new FileNotFoundException($"Template file not found: {templatePath}");

            _workbook = LoadWorkbook(templatePath);
            _designer = new WorkbookDesigner(_workbook);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading template: {ex.Message}");
            throw;
        }
    }

    // Wrapper for setting a variable (uses SetDataSource as fallback)
    public void SetVariable(string name, object value)
    {
        try
        {
            Console.WriteLine($"SetVariable called: Name = {name}, Value = {value}");
            // WorkbookDesigner.SetVariable may not be available in some versions;
            // SetDataSource works for scalar values as an alternative.
            _designer.SetDataSource(name, value);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error setting variable '{name}': {ex.Message}");
            throw;
        }
    }

    // Executes smart marker processing
    public void ProcessSmartMarkers()
    {
        try
        {
            _designer.Process();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing smart markers: {ex.Message}");
            throw;
        }
    }

    // Saves the workbook to the specified path
    public void Save(string outputPath)
    {
        try
        {
            // Ensure the directory exists
            string dir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            SaveWorkbook(_workbook, outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving workbook: {ex.Message}");
            throw;
        }
    }

    // -------------------------------------------------------------------------
    // Lifecycle rule placeholders – DO NOT modify the internal logic
    // -------------------------------------------------------------------------

    private Workbook CreateWorkbook()
    {
        // {CreateWorkbook} – creates a new empty workbook
        return new Workbook();
    }

    private Workbook LoadWorkbook(string path)
    {
        // {LoadWorkbook} – loads a workbook from the specified file
        return new Workbook(path);
    }

    private void SaveWorkbook(Workbook wb, string path)
    {
        // {SaveWorkbook} – saves the workbook to the specified file
        wb.Save(path);
    }
}

// Example usage
class Program
{
    static void Main()
    {
        try
        {
            var processor = new SmartMarkerProcessor();

            // Load a template that contains smart markers
            processor.LoadTemplate("Template.xlsx");

            // Set variables – each call will be logged
            processor.SetVariable("CustomerName", "John Doe");
            processor.SetVariable("OrderDate", DateTime.Now);
            processor.SetVariable("TotalAmount", 1234.56);

            // Process the smart markers with the variables set above
            processor.ProcessSmartMarkers();

            // Save the result
            processor.Save("Result.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unhandled exception: {ex.Message}");
        }
    }
}
