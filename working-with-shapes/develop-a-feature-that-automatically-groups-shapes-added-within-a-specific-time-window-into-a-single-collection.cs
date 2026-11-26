// Title: Automatically group worksheet shapes added within a specified time window using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# group shapes added within a few seconds | auto group worksheet shapes based on addition time Aspose.Cells | C# shape grouping after delay using Aspose.Cells Shape.Group | how to batch group shapes in Aspose.Cells by timestamp | group multiple shapes programmatically in Aspose.Cells .NET
// Tags: time‑window shape grouping Aspose.Cells | automatic grouping of worksheet shapes C# | using Shape.Group method Aspose.Cells | shape addition timestamp tracking Aspose.Cells | batch shape collection handling .NET

using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeAutoGroupingDemo
{
    // The example sets a configurable time window, records each newly added shape with its timestamp, and automatically groups shapes whose addition times fall within that window by calling Worksheet.Shapes.Group. Remaining pending shapes are forced into a final group before the workbook is saved.
    class Program
    {
        // Time window within which shapes will be grouped (e.g., 5 seconds)
        private static readonly TimeSpan GroupingWindow = TimeSpan.FromSeconds(5);

        // Holds shapes that have been added together with their addition time
        private static readonly List<(Shape shape, DateTime addedAt)> pendingShapes = new List<(Shape, DateTime)>();

        static void Main(string[] args)
        {
            try
            {
                // Initialize a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                ProgramHolder.Workbook = workbook; // store reference for helper methods
                Worksheet worksheet = workbook.Worksheets[0];
                ShapeCollection shapes = worksheet.Shapes;

                Console.WriteLine("Adding shapes...");

                // Add first rectangle
                AddShape(() => shapes.AddRectangle(2, 0, 2, 0, 80, 150));

                // Wait 2 seconds (still within the grouping window)
                Thread.Sleep(2000);

                // Add a textbox
                AddShape(() => shapes.AddTextBox(5, 0, 5, 0, 80, 150));

                // Wait 3 seconds (total 5 seconds from first shape, still within window)
                Thread.Sleep(3000);

                // Add a picture if a sample image file exists
                string imagePath = "sample.png";
                if (File.Exists(imagePath))
                {
                    using (FileStream fs = File.OpenRead(imagePath))
                    {
                        AddShape(() => shapes.AddPicture(8, 2, 0, 0, fs));
                    }
                }
                else
                {
                    Console.WriteLine($"Image file '{imagePath}' not found. Skipping picture addition.");
                }

                // Wait 6 seconds (outside the grouping window, next shapes will start a new group)
                Thread.Sleep(6000);

                // Add another rectangle – should start a new group
                AddShape(() => shapes.AddRectangle(12, 0, 12, 0, 80, 150));

                // Wait 1 second and add another shape – they will be grouped together
                Thread.Sleep(1000);
                AddShape(() => shapes.AddTextBox(15, 0, 15, 0, 80, 150));

                // Final attempt to group any remaining pending shapes
                TryGroupPendingShapes(forceGroup: true);

                // Save the workbook
                string outputPath = "ShapeAutoGroupingDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        /// <param name="addAction">An action that adds a shape to the ShapeCollection.</param>
        private static void AddShape(Action addAction)
        {
            try
            {
                // Execute the shape-adding action
                addAction.Invoke();

                // Retrieve the shape that was just added (last in the collection)
                Worksheet worksheet = GetCurrentWorksheet();
                if (worksheet.Shapes.Count == 0)
                {
                    Console.WriteLine("No shape was added.");
                    return;
                }

                Shape addedShape = worksheet.Shapes[worksheet.Shapes.Count - 1];

                // Record the shape with its addition timestamp
                pendingShapes.Add((addedShape, DateTime.Now));

                // Attempt to group shapes that satisfy the time window condition
                TryGroupPendingShapes();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to add shape: {ex.Message}");
            }
        }

        /// <summary>
        /// Attempts to group pending shapes that were added within the grouping window.
        /// If forceGroup is true, all pending shapes are grouped regardless of timing.
        /// </summary>
        /// <param name="forceGroup">When true, groups all pending shapes.</param>
        private static void TryGroupPendingShapes(bool forceGroup = false)
        {
            if (pendingShapes.Count < 2)
                return; // Need at least two shapes to form a group

            DateTime now = DateTime.Now;
            List<Shape> eligible = new List<Shape>();

            foreach (var (shape, addedAt) in pendingShapes)
            {
                if (forceGroup || (now - addedAt) <= GroupingWindow)
                {
                    eligible.Add(shape);
                }
            }

            if (eligible.Count >= 2)
            {
                try
                {
                    Worksheet worksheet = GetCurrentWorksheet();

                    // Group the shapes; the Group method returns a new Shape representing the group
                    Shape groupShape = worksheet.Shapes.Group(eligible.ToArray());

                    // Optionally, set a name for the group
                    groupShape.Name = $"AutoGroup_{DateTime.Now:yyyyMMdd_HHmmss}";

                    // Remove the grouped shapes from pending list
                    pendingShapes.RemoveAll(item => eligible.Contains(item.shape));

                    Console.WriteLine($"Grouped {eligible.Count} shapes into '{groupShape.Name}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to group shapes: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Retrieves the current worksheet from the active workbook.
        /// </summary>
        /// <returns>The active Worksheet.</returns>
        private static Worksheet GetCurrentWorksheet()
        {
            return ProgramHolder.Workbook.Worksheets[0];
        }

        /// <summary>
        /// Holds static references to objects that need to be accessed from helper methods.
        /// </summary>
        private static class ProgramHolder
        {
            public static Workbook Workbook;
        }
    }
}
