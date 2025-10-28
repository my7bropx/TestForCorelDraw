using System;
using System.Collections.Generic;
using System.Linq;

namespace CorelSmartFill
{
    /// <summary>
    /// Represents a single placed element in the container
    /// Stores position, rotation, and reference to the CorelDRAW shape
    /// </summary>
    public class PlacedElement
    {
        // Center X coordinate
        public double X { get; set; }

        // Center Y coordinate
        public double Y { get; set; }

        // Rotation angle in degrees (0-360)
        public double Rotation { get; set; }

        // Width of the element
        public double Width { get; set; }

        // Height of the element
        public double Height { get; set; }

        // Reference to the actual CorelDRAW shape object
        public dynamic? Shape { get; set; }

        /// <summary>
        /// Check if this element collides (overlaps) with another element
        /// Uses simple bounding box collision detection
        /// </summary>
        /// <param name="other">The other element to check against</param>
        /// <param name="margin">Extra space to add around elements (for spacing)</param>
        /// <returns>True if collision detected, false otherwise</returns>
        public bool CollidesWith(PlacedElement other, double margin = 0)
        {
            // Calculate distance between centers
            double dx = Math.Abs(this.X - other.X);
            double dy = Math.Abs(this.Y - other.Y);

            // Calculate minimum distance needed to avoid collision
            // We add both half-widths and both half-heights
            double minDx = (this.Width + other.Width) / 2 + margin;
            double minDy = (this.Height + other.Height) / 2 + margin;

            // Collision occurs if actual distance is less than minimum
            return dx < minDx && dy < minDy;
        }
    }

    /// <summary>
    /// Smart packing algorithm for filling container with elements
    /// Uses multiple strategies to achieve efficient space filling
    /// </summary>
    public class PackingAlgorithm
    {
        // Random number generator for random placement and rotation
        private Random random = new Random();

        // List of all placed elements (for collision detection)
        private List<PlacedElement> placedElements = new List<PlacedElement>();

        // Connection to CorelDRAW
        private CorelDRAWConnection corelConnection;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="connection">Active CorelDRAW connection</param>
        public PackingAlgorithm(CorelDRAWConnection connection)
        {
            this.corelConnection = connection;
        }

        /// <summary>
        /// Main packing function - fills container with elements
        /// </summary>
        /// <param name="container">The container shape to fill</param>
        /// <param name="fillElements">List of elements to use for filling</param>
        /// <param name="density">How many placement attempts (higher = more dense)</param>
        /// <param name="spacing">Minimum space between elements as percentage</param>
        /// <param name="allowRotation">Allow random rotation for better fit</param>
        /// <param name="smartRotation">Use smart rotation (90° increments) instead of random</param>
        /// <param name="preventOverlap">Check for collisions before placing</param>
        /// <returns>Number of elements successfully placed</returns>
        public int FillContainer(
            dynamic container,
            List<dynamic> fillElements,
            int density,
            double spacing,
            bool allowRotation,
            bool smartRotation,
            bool preventOverlap)
        {
            // STEP 1: Clear any previous placements
            placedElements.Clear();

            // STEP 2: Get container dimensions
            var bounds = corelConnection.GetBoundingBox(container);
            double leftX = bounds.leftX;
            double rightX = bounds.rightX;
            double topY = bounds.topY;
            double bottomY = bounds.bottomY;

            // Calculate container width and height
            double containerWidth = rightX - leftX;
            double containerHeight = topY - bottomY;

            // STEP 3: Calculate average element size
            double totalWidth = 0;
            double totalHeight = 0;
            foreach (var elem in fillElements)
            {
                totalWidth += elem.SizeWidth;
                totalHeight += elem.SizeHeight;
            }
            double avgWidth = totalWidth / fillElements.Count;
            double avgHeight = totalHeight / fillElements.Count;

            // STEP 4: Calculate spacing in actual units
            double spacingAmount = ((avgWidth + avgHeight) / 2) * spacing;

            // STEP 5: Enable optimization for faster performance
            corelConnection.SetOptimization(true);
            corelConnection.BeginUndoGroup("Smart Fill Container");

            // STEP 6: Track progress
            int placed = 0;
            int attempts = 0;
            int maxAttempts = density;

            // Current element index (cycle through fillElements)
            int elemIndex = 0;

            // STEP 7: Main placement loop
            while (attempts < maxAttempts && placed < maxAttempts)
            {
                attempts++;

                // Generate random position within container bounds
                // Add some margin to avoid placing too close to edges
                double margin = Math.Max(avgWidth, avgHeight) / 2;
                double x = leftX + margin + random.NextDouble() * (containerWidth - 2 * margin);
                double y = bottomY + margin + random.NextDouble() * (containerHeight - 2 * margin);

                // Check if point is inside container
                if (!corelConnection.IsPointInside(container, x, y))
                {
                    continue; // Try next position
                }

                // Determine rotation angle
                double rotation = 0;
                if (allowRotation)
                {
                    if (smartRotation)
                    {
                        // Smart rotation: only 0°, 90°, 180°, 270°
                        // This works better for rectangular elements
                        int[] smartAngles = { 0, 90, 180, 270 };
                        rotation = smartAngles[random.Next(smartAngles.Length)];
                    }
                    else
                    {
                        // Free rotation: any angle 0-360°
                        rotation = random.NextDouble() * 360;
                    }
                }

                // Get current fill element (cycle through list)
                dynamic sourceElement = fillElements[elemIndex];
                double elemWidth = sourceElement.SizeWidth;
                double elemHeight = sourceElement.SizeHeight;

                // Create a temporary PlacedElement to check collision
                PlacedElement newElement = new PlacedElement
                {
                    X = x,
                    Y = y,
                    Rotation = rotation,
                    Width = elemWidth,
                    Height = elemHeight
                };

                // STEP 8: Check for collisions if overlap prevention is enabled
                bool canPlace = true;

                if (preventOverlap)
                {
                    // Check against all previously placed elements
                    foreach (var existing in placedElements)
                    {
                        if (newElement.CollidesWith(existing, spacingAmount))
                        {
                            canPlace = false;
                            break;
                        }
                    }
                }

                // STEP 9: Additional boundary check
                // Make sure element doesn't extend outside container
                if (canPlace)
                {
                    // Check multiple points around the element
                    double checkRadius = Math.Max(elemWidth, elemHeight) / 2;
                    bool allPointsInside = true;

                    // Check 8 points around the element (like a circle)
                    for (int angle = 0; angle < 360; angle += 45)
                    {
                        double rad = angle * Math.PI / 180;
                        double checkX = x + Math.Cos(rad) * checkRadius;
                        double checkY = y + Math.Sin(rad) * checkRadius;

                        if (!corelConnection.IsPointInside(container, checkX, checkY))
                        {
                            allPointsInside = false;
                            break;
                        }
                    }

                    canPlace = allPointsInside;
                }

                // STEP 10: Place element if all checks passed
                if (canPlace)
                {
                    try
                    {
                        // Create actual shape in CorelDRAW
                        dynamic newShape = corelConnection.DuplicateAndPlace(
                            sourceElement, x, y, rotation);

                        // Store reference to placed element
                        newElement.Shape = newShape;
                        placedElements.Add(newElement);

                        placed++;

                        // Move to next element in the list
                        elemIndex = (elemIndex + 1) % fillElements.Count;
                    }
                    catch
                    {
                        // If placement failed, continue trying
                        continue;
                    }
                }
            }

            // STEP 11: Finish up
            corelConnection.EndUndoGroup();
            corelConnection.SetOptimization(false);
            corelConnection.Refresh();

            return placed;
        }

        /// <summary>
        /// Adjust density of already placed elements
        /// This removes or adds elements to achieve target density
        /// </summary>
        /// <param name="targetDensity">Target density as percentage (0-100)</param>
        public void AdjustDensity(double targetDensity)
        {
            // Calculate how many elements we should have
            int currentCount = placedElements.Count;
            int targetCount = (int)(currentCount * (targetDensity / 100.0));

            if (targetCount < currentCount)
            {
                // REMOVE elements (reduce density)
                int toRemove = currentCount - targetCount;

                // Start from end of list and remove
                for (int i = 0; i < toRemove && placedElements.Count > 0; i++)
                {
                    int lastIndex = placedElements.Count - 1;
                    var elem = placedElements[lastIndex];

                    try
                    {
                        // Delete shape from CorelDRAW
                        if (elem.Shape != null)
                        {
                            elem.Shape.Delete();
                        }
                    }
                    catch
                    {
                        // Ignore deletion errors
                    }

                    placedElements.RemoveAt(lastIndex);
                }
            }
            else if (targetCount > currentCount)
            {
                // ADD more elements (increase density)
                // Note: This would require re-running the packing algorithm
                // For now, we'll just note this limitation
                // In a full implementation, we'd call FillContainer again with adjusted parameters
            }

            // Refresh CorelDRAW display
            corelConnection.Refresh();
        }

        /// <summary>
        /// Get current statistics about placed elements
        /// </summary>
        /// <returns>Tuple of (count, coverage percentage)</returns>
        public (int count, double coverage) GetStatistics(dynamic container)
        {
            int count = placedElements.Count;

            // Calculate coverage (very rough estimate)
            var bounds = corelConnection.GetBoundingBox(container);
            double containerArea = (bounds.rightX - bounds.leftX) *
                                  (bounds.topY - bounds.bottomY);

            double elementsArea = 0;
            foreach (var elem in placedElements)
            {
                elementsArea += elem.Width * elem.Height;
            }

            double coverage = (elementsArea / containerArea) * 100;
            coverage = Math.Min(coverage, 100); // Cap at 100%

            return (count, coverage);
        }

        /// <summary>
        /// Clear all placed elements
        /// </summary>
        public void Clear()
        {
            placedElements.Clear();
        }

        /// <summary>
        /// Get list of all placed elements
        /// </summary>
        public List<PlacedElement> GetPlacedElements()
        {
            return new List<PlacedElement>(placedElements);
        }
    }
}
