using System;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;

namespace Deaxo.MakeParallel.Commands
{
    /// <summary>
    /// Analyzes elements to determine their direction vectors with comprehensive MEP support.
    /// Supports: Grids, Reference Planes, Lines, Family Instances, MEP elements, and Section views.
    /// </summary>
    public class ElementDirectionAnalyzer
    {
        private readonly Document _doc;

        public ElementDirectionAnalyzer(Document doc)
        {
            _doc = doc;
        }

        /// <summary>
        /// Gets the direction vector for any supported element type.
        /// </summary>
        /// <param name="element">Element to analyze</param>
        /// <returns>Direction vector or null if unsupported</returns>
        public XYZ GetDirection(Element element)
        {
            if (element == null) return null;

            try
            {
                // Try each direction detection method in priority order
                return GetGridDirection(element) ??
                       GetReferencePlaneDirection(element) ??
                       GetMepDirection(element) ??
                       GetFamilyInstanceDirection(element) ??
                       GetLineBasedDirection(element) ??
                       GetSectionDirection(element);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Gets the origin point for any supported element type.
        /// </summary>
        /// <param name="element">Element to analyze</param>
        /// <returns>Origin point or null if unsupported</returns>
        public XYZ GetOrigin(Element element)
        {
            if (element == null) return null;

            try
            {
                return GetGridOrigin(element) ??
                       GetReferencePlaneOrigin(element) ??
                       GetMepOrigin(element) ??
                       GetFamilyInstanceOrigin(element) ??
                       GetLineBasedOrigin(element) ??
                       GetSectionOrigin(element);
            }
            catch
            {
                return null;
            }
        }

        #region Direction Detection Methods

        private XYZ GetGridDirection(Element element)
        {
            if (element is Grid grid)
            {
                var curve = grid.Curve;
                return CalculateDirection(curve);
            }
            return null;
        }

        private XYZ GetReferencePlaneDirection(Element element)
        {
            if (element is ReferencePlane refPlane)
                return refPlane.Direction;
            return null;
        }

        private XYZ GetMepDirection(Element element)
        {
            // Pipes
            if (element is Pipe pipe)
            {
                var locationCurve = pipe.Location as LocationCurve;
                return locationCurve != null ? CalculateDirection(locationCurve.Curve) : null;
            }

            // Ducts
            if (element is Duct duct)
            {
                var locationCurve = duct.Location as LocationCurve;
                return locationCurve != null ? CalculateDirection(locationCurve.Curve) : null;
            }

            // Cable Trays
            if (element is CableTray cableTray)
            {
                var locationCurve = cableTray.Location as LocationCurve;
                return locationCurve != null ? CalculateDirection(locationCurve.Curve) : null;
            }

            // Conduits
            if (element is Conduit conduit)
            {
                var locationCurve = conduit.Location as LocationCurve;
                return locationCurve != null ? CalculateDirection(locationCurve.Curve) : null;
            }

            // MEP Curve-based elements (flexible ducts, pipes, etc.)
            if (element.Location is LocationCurve mepLocationCurve)
            {
                // Check if it's an MEP element by category
                var category = element.Category;
                if (category != null)
                {
                    var catId = category.Id.IntegerValue;
                    if (IsMepCategory(catId))
                    {
                        return CalculateDirection(mepLocationCurve.Curve);
                    }
                }
            }

            return null;
        }

        private XYZ GetFamilyInstanceDirection(Element element)
        {
            if (element is FamilyInstance familyInstance)
            {
                // Try facing orientation first (for hosted elements like doors, windows)
                try
                {
                    return familyInstance.FacingOrientation;
                }
                catch
                {
                    // If facing orientation fails, try transform basis X
                    try
                    {
                        var transform = familyInstance.GetTransform();
                        return transform.BasisX;
                    }
                    catch
                    {
                        return null;
                    }
                }
            }
            return null;
        }

        private XYZ GetLineBasedDirection(Element element)
        {
            if (element.Location is LocationCurve locationCurve)
            {
                return CalculateDirection(locationCurve.Curve);
            }
            return null;
        }

        private XYZ GetSectionDirection(Element element)
        {
            try
            {
                var sketchParam = element.get_Parameter(BuiltInParameter.VIEW_FIXED_SKETCH_PLANE);
                if (sketchParam != null && sketchParam.HasValue)
                {
                    var sketchPlane = _doc.GetElement(sketchParam.AsElementId()) as SketchPlane;
                    if (sketchPlane != null)
                    {
                        var view = _doc.GetElement(sketchPlane.OwnerViewId) as ViewSection;
                        return view?.RightDirection;
                    }
                }
            }
            catch { }
            return null;
        }

        #endregion

        #region Origin Detection Methods

        private XYZ GetGridOrigin(Element element)
        {
            if (element is Grid grid)
            {
                var curve = grid.Curve;
                return CalculateOrigin(curve);
            }
            return null;
        }

        private XYZ GetReferencePlaneOrigin(Element element)
        {
            if (element is ReferencePlane refPlane)
                return refPlane.GetPlane().Origin;
            return null;
        }

        private XYZ GetMepOrigin(Element element)
        {
            // For MEP elements, get the midpoint of the curve
            if (element is Pipe || element is Duct || element is CableTray || element is Conduit)
            {
                if (element.Location is LocationCurve locationCurve)
                {
                    var curve = locationCurve.Curve;
                    return curve.Evaluate(0.5, true); // Midpoint
                }
            }
            return null;
        }

        private XYZ GetFamilyInstanceOrigin(Element element)
        {
            if (element is FamilyInstance familyInstance)
            {
                var transform = familyInstance.GetTransform();
                return transform.Origin;
            }
            return null;
        }

        private XYZ GetLineBasedOrigin(Element element)
        {
            if (element.Location is LocationCurve locationCurve)
            {
                return CalculateOrigin(locationCurve.Curve);
            }
            return null;
        }

        private XYZ GetSectionOrigin(Element element)
        {
            try
            {
                var sketchParam = element.get_Parameter(BuiltInParameter.VIEW_FIXED_SKETCH_PLANE);
                if (sketchParam != null && sketchParam.HasValue)
                {
                    var sketchPlane = _doc.GetElement(sketchParam.AsElementId()) as SketchPlane;
                    if (sketchPlane != null)
                    {
                        var view = _doc.GetElement(sketchPlane.OwnerViewId) as ViewSection;
                        return view?.Origin;
                    }
                }
            }
            catch { }
            return null;
        }

        #endregion

        #region Special Case Handling

        /// <summary>
        /// Checks if an element is an elevation view.
        /// </summary>
        public bool IsElevationView(Element element)
        {
            try
            {
                var sketchParam = element.get_Parameter(BuiltInParameter.VIEW_FIXED_SKETCH_PLANE);
                if (sketchParam != null && sketchParam.HasValue)
                {
                    var sketchPlane = _doc.GetElement(sketchParam.AsElementId()) as SketchPlane;
                    if (sketchPlane != null)
                    {
                        var view = _doc.GetElement(sketchPlane.OwnerViewId) as View;
                        return view?.ViewType == ViewType.Elevation;
                    }
                }
            }
            catch { }
            return false;
        }

        /// <summary>
        /// Gets the elevation marker associated with an elevation view.
        /// </summary>
        public ElevationMarker GetElevationMarker(Element element)
        {
            try
            {
                var sketchParam = element.get_Parameter(BuiltInParameter.VIEW_FIXED_SKETCH_PLANE);
                if (sketchParam != null && sketchParam.HasValue)
                {
                    var sketchPlane = _doc.GetElement(sketchParam.AsElementId()) as SketchPlane;
                    if (sketchPlane != null)
                    {
                        var view = _doc.GetElement(sketchPlane.OwnerViewId) as View;
                        if (view != null)
                        {
                            var elevationMarkers = new FilteredElementCollector(_doc)
                                .OfClass(typeof(ElevationMarker))
                                .Cast<ElevationMarker>();

                            foreach (var marker in elevationMarkers)
                            {
                                for (int i = 0; i < 4; i++)
                                {
                                    var viewId = marker.GetViewId(i);
                                    if (viewId == view.Id)
                                        return marker;
                                }
                            }
                        }
                    }
                }
            }
            catch { }
            return null;
        }

        #endregion

        #region Helper Methods

        /// <summary>
        /// Determines if a category ID represents an MEP element.
        /// </summary>
        private bool IsMepCategory(int categoryId)
        {
            var mepCategories = new[]
            {
                (int)BuiltInCategory.OST_PipeCurves,
                (int)BuiltInCategory.OST_DuctCurves,
                (int)BuiltInCategory.OST_CableTray,
                (int)BuiltInCategory.OST_Conduit,
                (int)BuiltInCategory.OST_FlexPipeCurves,
                (int)BuiltInCategory.OST_FlexDuctCurves,
                (int)BuiltInCategory.OST_PipeAccessory,
                (int)BuiltInCategory.OST_PipeFitting,
                (int)BuiltInCategory.OST_DuctAccessory,
                (int)BuiltInCategory.OST_DuctFitting,
                (int)BuiltInCategory.OST_CableTrayFitting,
                (int)BuiltInCategory.OST_ConduitFitting
            };

            return mepCategories.Contains(categoryId);
        }

        /// <summary>
        /// Calculates direction vector from curve endpoints.
        /// </summary>
        private XYZ CalculateDirection(Curve curve)
        {
            if (curve == null) return null;

            try
            {
                var startPoint = curve.GetEndPoint(0);
                var endPoint = curve.GetEndPoint(1);
                return (endPoint - startPoint).Normalize();
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Calculates origin point from curve (midpoint or start point).
        /// </summary>
        private XYZ CalculateOrigin(Curve curve)
        {
            if (curve == null) return null;

            try
            {
                // Try to get midpoint first
                return curve.Evaluate(0.5, true);
            }
            catch
            {
                try
                {
                    // Fallback to start point
                    return curve.GetEndPoint(0);
                }
                catch
                {
                    return null;
                }
            }
        }

        #endregion
    }
}