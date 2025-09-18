using System;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Deaxo.MakeParallel.Commands;

namespace Deaxo.MakeParallel.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class MakeParallelCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // Step 1: Select reference element (no dialog - just status bar)
                Reference reference1;
                Element element1;
                try
                {
                    reference1 = uidoc.Selection.PickObject(ObjectType.Element, "Pick reference element (first element)");
                    element1 = doc.GetElement(reference1);
                }
                catch (OperationCanceledException)
                {
                    return Result.Cancelled;
                }

                if (element1 == null)
                {
                    TaskDialog.Show("Error", "Failed to get reference element.");
                    return Result.Failed;
                }

                // Step 2: Select target element (no dialog - just status bar)
                Reference reference2;
                Element element2;
                try
                {
                    reference2 = uidoc.Selection.PickObject(ObjectType.Element, "Pick target element (element to rotate)");
                    element2 = doc.GetElement(reference2);
                }
                catch (OperationCanceledException)
                {
                    return Result.Cancelled;
                }

                if (element2 == null)
                {
                    TaskDialog.Show("Error", "Failed to get target element.");
                    return Result.Failed;
                }

                // Step 3: Get directions and perform rotation in one transaction
                var analyzer = new ElementDirectionAnalyzer(doc);

                var direction1 = analyzer.GetDirection(element1);
                var direction2 = analyzer.GetDirection(element2);

                if (direction1 == null || direction2 == null)
                {
                    TaskDialog.Show("Unsupported Element",
                        "Cannot determine direction for one or both selected elements.\n\n" +
                        "Supported elements: Grids, Reference Planes, Lines, Family Instances, " +
                        "MEP elements (pipes, ducts, cable trays, conduits), and Section views.");
                    return Result.Failed;
                }

                // Step 4: Calculate and perform rotation
                var calculator = new ParallelCalculator();
                var rotationData = calculator.CalculateRotation(direction1, direction2, element2);

                if (rotationData.Angle == 0)
                {
                    TaskDialog.Show("DEAXO - Make Parallel", "Elements are already parallel!");
                    return Result.Succeeded;
                }

                // Handle special cases (elevation markers)
                Element elementToRotate = element2;
                if (analyzer.IsElevationView(element2))
                {
                    var elevationMarker = analyzer.GetElevationMarker(element2);
                    if (elevationMarker != null)
                    {
                        elementToRotate = elevationMarker;
                    }
                }

                // Perform rotation
                using (Transaction t = new Transaction(doc, "DEAXO - Make Parallel"))
                {
                    t.Start();

                    try
                    {
                        if (elementToRotate.Location != null)
                        {
                            elementToRotate.Location.Rotate(rotationData.Axis, rotationData.Angle);

                            // Single success message with rotation info
                            TaskDialog.Show("DEAXO - Make Parallel",
                                $"Elements made parallel.\nRotation: {Math.Abs(rotationData.Angle * 180 / Math.PI):F1}°");
                        }
                        else
                        {
                            TaskDialog.Show("Error", "Cannot rotate this element - no location property.");
                            t.RollBack();
                            return Result.Failed;
                        }
                    }
                    catch (Exception ex)
                    {
                        TaskDialog.Show("Rotation Failed",
                            $"Could not rotate element: {ex.Message}\n\nElement may be constrained or locked.");
                        t.RollBack();
                        return Result.Failed;
                    }

                    t.Commit();
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show("Error", $"An error occurred: {ex.Message}");
                return Result.Failed;
            }
        }
    }
}