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
                // Step 1: Select reference element
                TaskDialog.Show("DEAXO - Make Parallel", "Pick reference element (first element)");

                Reference reference1;
                Element element1;
                try
                {
                    reference1 = uidoc.Selection.PickObject(ObjectType.Element, "Pick reference element");
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

                // Step 2: Select target element
                TaskDialog.Show("DEAXO - Make Parallel", "Pick target element (element to rotate)");

                Reference reference2;
                Element element2;
                try
                {
                    reference2 = uidoc.Selection.PickObject(ObjectType.Element, "Pick target element");
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

                // Step 3: Get directions using enhanced element analysis
                var analyzer = new ElementDirectionAnalyzer(doc);

                var direction1 = analyzer.GetDirection(element1);
                var direction2 = analyzer.GetDirection(element2);

                if (direction1 == null)
                {
                    TaskDialog.Show("Unsupported Element",
                        $"Cannot determine direction for reference element of type '{element1.GetType().Name}'. " +
                        "Supported elements include: Grids, Reference Planes, Lines, Family Instances, " +
                        "MEP elements (pipes, ducts, cable trays, conduits), and Section views.");
                    return Result.Failed;
                }

                if (direction2 == null)
                {
                    TaskDialog.Show("Unsupported Element",
                        $"Cannot determine direction for target element of type '{element2.GetType().Name}'. " +
                        "Supported elements include: Grids, Reference Planes, Lines, Family Instances, " +
                        "MEP elements (pipes, ducts, cable trays, conduits), and Section views.");
                    return Result.Failed;
                }

                // Step 4: Calculate angle and rotation axis
                var calculator = new ParallelCalculator();
                var rotationData = calculator.CalculateRotation(direction1, direction2, element2);

                if (rotationData.Angle == 0)
                {
                    TaskDialog.Show("DEAXO - Make Parallel", "Elements are already parallel!");
                    return Result.Succeeded;
                }

                // Step 5: Handle special cases (elevation markers)
                Element elementToRotate = element2;
                if (analyzer.IsElevationView(element2))
                {
                    var elevationMarker = analyzer.GetElevationMarker(element2);
                    if (elevationMarker != null)
                    {
                        elementToRotate = elevationMarker;
                    }
                }

                // Step 6: Perform rotation
                using (Transaction t = new Transaction(doc, "DEAXO - Make Parallel"))
                {
                    t.Start();

                    try
                    {
                        if (elementToRotate.Location != null)
                        {
                            elementToRotate.Location.Rotate(rotationData.Axis, rotationData.Angle);

                            TaskDialog.Show("DEAXO - Make Parallel",
                                $"Successfully made elements parallel.\n" +
                                $"Rotation angle: {Math.Abs(rotationData.Angle * 180 / Math.PI):F2}°");
                        }
                        else
                        {
                            TaskDialog.Show("Error",
                                "Cannot rotate this element - it doesn't have a location property. " +
                                "This might be a system family or non-moveable element.");
                            t.RollBack();
                            return Result.Failed;
                        }
                    }
                    catch (Exception ex)
                    {
                        TaskDialog.Show("Rotation Error",
                            $"Failed to rotate element: {ex.Message}\n\n" +
                            "This element might be constrained, locked, or not allowed to rotate.");
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
                TaskDialog.Show("DEAXO - Make Parallel Error", $"An error occurred: {ex.Message}");
                return Result.Failed;
            }
        }
    }
}