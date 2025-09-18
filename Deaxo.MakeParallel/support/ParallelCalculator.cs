using System;
using Autodesk.Revit.DB;

namespace Deaxo.MakeParallel.Commands
{
    /// <summary>
    /// Data container for rotation calculations.
    /// </summary>
    public class RotationData
    {
        public double Angle { get; set; }
        public Line Axis { get; set; }
    }

    /// <summary>
    /// Calculates the rotation needed to make two elements parallel in the XY plane.
    /// Mirrors the Python parallel calculation logic.
    /// </summary>
    public class ParallelCalculator
    {
        /// <summary>
        /// Calculates the rotation angle and axis needed to make element2 parallel to element1.
        /// </summary>
        /// <param name="direction1">Direction vector of reference element</param>
        /// <param name="direction2">Direction vector of target element</param>
        /// <param name="element2">Target element for origin calculation</param>
        /// <returns>Rotation data with angle and axis</returns>
        public RotationData CalculateRotation(XYZ direction1, XYZ direction2, Element element2)
        {
            if (direction1 == null || direction2 == null || element2 == null)
            {
                return new RotationData { Angle = 0, Axis = null };
            }

            // Project vectors to XY plane (set Z to 0)
            var xyV1 = new XYZ(direction1.X, direction1.Y, 0);
            var xyV2 = new XYZ(direction2.X, direction2.Y, 0);

            // Normalize vectors to ensure accurate angle calculation
            try
            {
                xyV1 = xyV1.Normalize();
                xyV2 = xyV2.Normalize();
            }
            catch
            {
                // If normalization fails, vectors might be zero-length
                return new RotationData { Angle = 0, Axis = null };
            }

            // Calculate angle between vectors
            double angle = xyV2.AngleTo(xyV1);

            // Adjust angle if greater than 90 degrees (π/2)
            // This ensures we take the smaller rotation angle
            if (angle > Math.PI / 2)
            {
                angle = angle - Math.PI;
            }

            // If angle is effectively zero, elements are already parallel
            if (Math.Abs(angle) < 0.001) // ~0.06 degrees tolerance
            {
                return new RotationData { Angle = 0, Axis = null };
            }

            // Calculate cross product to determine rotation axis
            var normal = xyV2.CrossProduct(xyV1);

            // Get origin of target element for rotation axis
            var analyzer = new ElementDirectionAnalyzer(element2.Document);
            var origin = analyzer.GetOrigin(element2);

            if (origin == null)
            {
                // Fallback: use element's bounding box center
                var bb = element2.get_BoundingBox(null);
                if (bb != null)
                {
                    origin = (bb.Min + bb.Max) / 2;
                }
                else
                {
                    origin = XYZ.Zero;
                }
            }

            // Create rotation axis line
            var axisEnd = origin + normal;
            var axis = Line.CreateBound(origin, axisEnd);

            return new RotationData
            {
                Angle = angle,
                Axis = axis
            };
        }

        /// <summary>
        /// Checks if two direction vectors are already parallel within tolerance.
        /// </summary>
        /// <param name="direction1">First direction vector</param>
        /// <param name="direction2">Second direction vector</param>
        /// <param name="toleranceDegrees">Tolerance in degrees (default 0.1°)</param>
        /// <returns>True if vectors are parallel within tolerance</returns>
        public bool AreVectorsParallel(XYZ direction1, XYZ direction2, double toleranceDegrees = 0.1)
        {
            if (direction1 == null || direction2 == null) return false;

            try
            {
                // Project to XY plane
                var xyV1 = new XYZ(direction1.X, direction1.Y, 0).Normalize();
                var xyV2 = new XYZ(direction2.X, direction2.Y, 0).Normalize();

                // Calculate angle
                double angle = xyV2.AngleTo(xyV1);

                // Check if parallel (0°) or anti-parallel (180°)
                double toleranceRad = toleranceDegrees * Math.PI / 180;

                return (angle < toleranceRad) ||
                       (Math.Abs(angle - Math.PI) < toleranceRad);
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Converts angle from radians to degrees for display purposes.
        /// </summary>
        /// <param name="angleRadians">Angle in radians</param>
        /// <returns>Angle in degrees</returns>
        public double RadiansToDegrees(double angleRadians)
        {
            return angleRadians * 180.0 / Math.PI;
        }

        /// <summary>
        /// Gets a user-friendly description of the rotation.
        /// </summary>
        /// <param name="rotationData">Rotation data</param>
        /// <returns>Human-readable description</returns>
        public string GetRotationDescription(RotationData rotationData)
        {
            if (rotationData?.Angle == null || Math.Abs(rotationData.Angle) < 0.001)
            {
                return "Elements are already parallel";
            }

            double degrees = Math.Abs(RadiansToDegrees(rotationData.Angle));
            string direction = rotationData.Angle > 0 ? "counterclockwise" : "clockwise";

            return $"Rotating {degrees:F2}° {direction}";
        }
    }
}