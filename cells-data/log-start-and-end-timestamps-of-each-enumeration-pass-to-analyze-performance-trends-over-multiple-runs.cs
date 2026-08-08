// Title: Log enumeration timestamps for worksheet rows using Aspose.Cells for .NET
// Description: Creates a workbook, fills five rows, then runs three GetEnumerator passes (normal, reversed, synchronized) over worksheet rows. Each pass logs start and end times and prints the duration, enabling performance comparison of different enumerator configurations.
// Keywords: Aspose.Cells row enumeration | enumeration timing .NET | worksheet performance logging | GetEnumerator reverse order | synchronized enumerator Aspose.Cells | C# performance measurement
// Common Searches: measure row enumeration time Aspose.Cells | log start and end timestamps GetEnumerator Aspose.Cells | performance comparison of Aspose.Cells enumerators | how to time worksheet row iteration in C#
// Developer Intent: Capture start and end timestamps for each worksheet row enumeration to evaluate the speed of normal, reversed, and synchronized GetEnumerator configurations.
// Use Cases: Establish a baseline execution time for normal‑order, non‑synchronized row enumeration. | Compare the speed of reversed‑order enumeration against the baseline. | Determine the overhead introduced by a synchronized enumerator in normal order.
// AI Prompts: Generate C# code that logs start and end times for each worksheet row enumeration using Aspose.Cells, including reverse and synchronized options. | Explain how to analyze the duration output to compare performance of different GetEnumerator settings in Aspose.Cells. | Suggest best practices for minimizing overhead when timing row enumeration in Aspose.Cells.

using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsPerformanceLog
{
    // Creates a workbook, fills five rows, then runs three GetEnumerator passes (normal, reversed, synchronized) over worksheet rows. Each pass logs start and end times and prints the duration, enabling performance comparison of different enumerator configurations.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some sample data in the first 5 rows
            for (int i = 0; i < 5; i++)
            {
                Row row = worksheet.Cells.Rows[i];
                row[0].PutValue($"Row {i + 1}");
                row[1].PutValue(i * 10);
            }

            // First enumeration pass: normal order, non‑synchronized enumerator
            LogEnumerationPass(
                () => worksheet.Cells.Rows.GetEnumerator(),
                "Normal order (non‑synchronized)");

            // Second enumeration pass: reversed order, non‑synchronized enumerator
            LogEnumerationPass(
                () => worksheet.Cells.Rows.GetEnumerator(true, false),
                "Reversed order (non‑synchronized)");

            // Third enumeration pass: normal order, synchronized enumerator
            LogEnumerationPass(
                () => worksheet.Cells.Rows.GetEnumerator(false, true),
                "Normal order (synchronized)");

            // Save the workbook (using standard save logic)
            workbook.Save("PerformanceLogDemo.xlsx");
        }

        /// <param name="enumeratorFactory">A function that returns a fresh IEnumerator for the collection.</param>
        /// <param name="description">Description of the enumeration pass.</param>
        private static void LogEnumerationPass(Func<IEnumerator> enumeratorFactory, string description)
        {
            Console.WriteLine($"--- {description} ---");

            DateTime startTime = DateTime.Now;
            Console.WriteLine($"Start: {startTime:O}");

            IEnumerator enumerator = enumeratorFactory();

            while (enumerator.MoveNext())
            {
                // Access the current element to ensure the enumerator advances.
                // In this demo we simply cast to Row; real logic can be placed here.
                Row currentRow = enumerator.Current as Row;
                // Optional: perform lightweight operation to avoid compiler optimizations removing the loop.
                if (currentRow != null && currentRow.Index >= 0) { /* no-op */ }
            }

            DateTime endTime = DateTime.Now;
            Console.WriteLine($"End:   {endTime:O}");
            Console.WriteLine($"Duration: {(endTime - startTime).TotalMilliseconds} ms");
            Console.WriteLine();
        }
    }
}
