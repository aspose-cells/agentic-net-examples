// Title: Save Processed Smart Markers Workbook as XLSX with Aspose.Cells for .NET
// Description: Shows how to load a smart‑marker template, bind an ArrayList data source, process the markers using WorkbookDesigner, and save the workbook as an XLSX file that keeps all generated content and formatting.
// Keywords: Aspose.Cells | C# | .NET | WorkbookDesigner | Smart Markers | Save as XLSX | Preserve formatting | Excel template | ArrayList data source | Export processed workbook
// Common Searches: Aspose.Cells save workbook after smart marker processing | How to keep smart marker formatting when exporting to XLSX | WorkbookDesigner process smart markers C# example | Export smart marker populated Excel as XLSX | Preserve styles after Aspose.Cells smart marker run
// Developer Intent: Generate a final XLSX file from a smart‑marker template while retaining all inserted data and cell styles.
// Use Cases: Create personalized employee reports from a smart‑marker template and export them as styled XLSX files. | Produce data‑driven invoices using smart markers, then save the finished documents without losing formatting. | Automate batch processing of multiple templates with different data sets, applying smart markers and storing each result as a formatted XLSX workbook.
// AI Prompts: Write C# code that loads an Excel template with smart markers, sets an ArrayList as the data source, processes the markers with WorkbookDesigner, and saves the output as XLSX preserving all styles. | Explain step‑by‑step how Aspose.Cells WorkbookDesigner populates a smart‑marker template and ensures the saved XLSX retains generated content and formatting. | Provide a concise guide for processing smart markers in a .NET workbook and exporting the result to XLSX using Aspose.Cells.

using System;
using System.Collections;
using Aspose.Cells;

// Shows how to load a smart‑marker template, bind an ArrayList data source, process the markers using WorkbookDesigner, and save the workbook as an XLSX file that keeps all generated content and formatting.
class Program
{
    static void Main()
    {
        // Load the workbook that contains smart markers (template)
        Workbook workbook = new Workbook("TemplateWithSmartMarkers.xlsx");

        // Create a WorkbookDesigner and associate it with the loaded workbook
        WorkbookDesigner designer = new WorkbookDesigner(workbook);

        // Prepare a sample data source (replace with actual data as needed)
        ArrayList persons = new ArrayList();
        persons.Add(new Person { Name = "John Doe", Age = 30 });
        persons.Add(new Person { Name = "Jane Smith", Age = 28 });

        // Set the data source for the smart markers
        designer.SetDataSource("Persons", persons);

        // Process the smart markers – this populates the workbook with data
        designer.Process();

        // Save the processed workbook as XLSX to preserve all generated content and formatting
        workbook.Save("ProcessedOutput.xlsx", SaveFormat.Xlsx);
    }
}

// Sample data class used in the data source (customize as required)
public class Person
{
    public string Name { get; set; }
    public int Age { get; set; }
}
