using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;

namespace CorelSmartFill
{
    /// <summary>
    /// Handles all communication with CorelDRAW via COM automation
    /// COM = Component Object Model - Windows technology for inter-process communication
    /// This allows our C# program to control CorelDRAW
    /// </summary>
    public class CorelDRAWConnection
    {
        // Store reference to CorelDRAW application
        // "dynamic" type means we don't need to know the exact type at compile time
        // This is necessary because different CorelDRAW versions have different COM interfaces
        private dynamic? corelApp = null;

        // Store reference to the currently active document
        private dynamic? activeDocument = null;

        /// <summary>
        /// Check if we're currently connected to CorelDRAW
        /// </summary>
        public bool IsConnected => corelApp != null;

        /// <summary>
        /// Connect to a running instance of CorelDRAW
        /// This looks for CorelDRAW that's already open on the computer
        /// </summary>
        /// <returns>True if connection successful, false otherwise</returns>
        public bool Connect()
        {
            try
            {
                // STEP 1: Try to get running instance of CorelDRAW
                // "CorelDRAW.Application" is the COM ProgID for CorelDRAW
                // Different versions:
                //   - CorelDRAW.Application.24 for X8/2018
                //   - CorelDRAW.Application.25 for 2019
                //   - etc.
                // We try the generic one first, which should work for most versions

                try
                {
                    // Try to connect to already running CorelDRAW
                    corelApp = Marshal.GetActiveObject("CorelDRAW.Application");
                }
                catch
                {
                    // If no instance running, try to start CorelDRAW
                    // Type.GetTypeFromProgID gets the COM type for CorelDRAW
                    Type? corelType = Type.GetTypeFromProgID("CorelDRAW.Application");

                    if (corelType == null)
                    {
                        throw new Exception("CorelDRAW is not installed or not registered for COM automation.");
                    }

                    // Activator.CreateInstance starts a new CorelDRAW instance
                    corelApp = Activator.CreateInstance(corelType);

                    // Make CorelDRAW visible to user
                    corelApp.Visible = true;
                }

                // STEP 2: Get the active document
                // The document is where all the shapes are stored
                activeDocument = corelApp.ActiveDocument;

                if (activeDocument == null)
                {
                    throw new Exception("No active document in CorelDRAW. Please open or create a document.");
                }

                return true;
            }
            catch (Exception ex)
            {
                // If anything went wrong, store error and return false
                throw new Exception($"Failed to connect to CorelDRAW: {ex.Message}");
            }
        }

        /// <summary>
        /// Get information about the selected shapes in CorelDRAW
        /// Returns: Container shape and list of fill elements
        /// </summary>
        /// <returns>Tuple of (container, fillElements)</returns>
        public (dynamic container, List<dynamic> fillElements) GetSelectedShapes()
        {
            if (!IsConnected || activeDocument == null)
            {
                throw new Exception("Not connected to CorelDRAW");
            }

            try
            {
                // STEP 1: Get all selected shapes
                // SelectionRange is a collection of all currently selected shapes
                dynamic selection = activeDocument.SelectionRange;

                // Check that at least 2 shapes are selected
                int count = selection.Count;
                if (count < 2)
                {
                    throw new Exception("Please select at least 2 objects: container curve and fill element(s)");
                }

                // STEP 2: Find the largest shape - this will be our container
                dynamic? container = null;
                double largestArea = 0;

                // Loop through all selected shapes
                for (int i = 1; i <= count; i++)
                {
                    dynamic shape = selection[i];

                    // Calculate area (width × height)
                    double width = shape.SizeWidth;
                    double height = shape.SizeHeight;
                    double area = width * height;

                    // If this is the largest so far, remember it
                    if (area > largestArea)
                    {
                        largestArea = area;
                        container = shape;
                    }
                }

                // STEP 3: Collect all other shapes as fill elements
                List<dynamic> fillElements = new List<dynamic>();

                for (int i = 1; i <= count; i++)
                {
                    dynamic shape = selection[i];

                    // Add to fill elements if it's not the container
                    if (shape != container)
                    {
                        fillElements.Add(shape);
                    }
                }

                if (fillElements.Count == 0)
                {
                    throw new Exception("No fill elements found. Make sure container is larger than fill elements.");
                }

                // STEP 4: Convert container to curves if needed
                // Curves are the most flexible shape type in CorelDRAW
                // cdrCurveShape = 16 (numeric value for curve type)
                if (container.Type != 16)
                {
                    container.ConvertToCurves();
                }

                return (container, fillElements);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error getting selected shapes: {ex.Message}");
            }
        }

        /// <summary>
        /// Check if a point (x, y) is inside the container curve
        /// This is crucial for boundary detection
        /// </summary>
        /// <param name="container">The container shape</param>
        /// <param name="x">X coordinate to test</param>
        /// <param name="y">Y coordinate to test</param>
        /// <returns>True if point is inside, false otherwise</returns>
        public bool IsPointInside(dynamic container, double x, double y)
        {
            try
            {
                // CorelDRAW's Curve object has IsPointInside method
                // This uses ray-casting algorithm internally
                return container.Curve.IsPointInside(x, y);
            }
            catch
            {
                // If method fails, return false (safer to assume outside)
                return false;
            }
        }

        /// <summary>
        /// Create a duplicate of a shape and place it at specified position and rotation
        /// </summary>
        /// <param name="sourceShape">The shape to duplicate</param>
        /// <param name="x">X coordinate for center</param>
        /// <param name="y">Y coordinate for center</param>
        /// <param name="rotation">Rotation angle in degrees</param>
        /// <returns>The newly created shape</returns>
        public dynamic DuplicateAndPlace(dynamic sourceShape, double x, double y, double rotation)
        {
            try
            {
                // STEP 1: Create duplicate
                // Duplicate() creates an exact copy with all properties preserved
                dynamic newShape = sourceShape.Duplicate();

                // STEP 2: Position the duplicate
                // CenterX/CenterY move the shape's center to specified coordinates
                newShape.CenterX = x;
                newShape.CenterY = y;

                // STEP 3: Rotate if needed
                // Rotate() rotates shape around its center
                // Angle is in degrees (0-360)
                if (rotation != 0)
                {
                    newShape.Rotate(rotation);
                }

                return newShape;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error duplicating shape: {ex.Message}");
            }
        }

        /// <summary>
        /// Get the bounding box of the container
        /// Returns the rectangular area that fully contains the shape
        /// </summary>
        /// <param name="container">The container shape</param>
        /// <returns>Tuple of (leftX, rightX, topY, bottomY)</returns>
        public (double leftX, double rightX, double topY, double bottomY) GetBoundingBox(dynamic container)
        {
            try
            {
                // CorelDRAW shapes have these properties built-in
                // LeftX: X coordinate of left edge
                // RightX: X coordinate of right edge
                // TopY: Y coordinate of top edge (higher value)
                // BottomY: Y coordinate of bottom edge (lower value)
                // NOTE: In CorelDRAW, Y increases upward (not downward like screen coordinates)

                double leftX = container.LeftX;
                double rightX = container.RightX;
                double topY = container.TopY;
                double bottomY = container.BottomY;

                return (leftX, rightX, topY, bottomY);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error getting bounding box: {ex.Message}");
            }
        }

        /// <summary>
        /// Begin a command group for undo
        /// This makes all operations between Begin and End a single undo step
        /// </summary>
        /// <param name="name">Name of the operation (shows in Edit > Undo menu)</param>
        public void BeginUndoGroup(string name)
        {
            if (activeDocument != null)
            {
                activeDocument.BeginCommandGroup(name);
            }
        }

        /// <summary>
        /// End the command group for undo
        /// </summary>
        public void EndUndoGroup()
        {
            if (activeDocument != null)
            {
                activeDocument.EndCommandGroup();
            }
        }

        /// <summary>
        /// Refresh CorelDRAW window to show changes
        /// </summary>
        public void Refresh()
        {
            if (corelApp != null)
            {
                try
                {
                    corelApp.ActiveWindow.Refresh();
                }
                catch
                {
                    // Ignore refresh errors
                }
            }
        }

        /// <summary>
        /// Enable/disable optimization mode
        /// When enabled, CorelDRAW doesn't redraw screen during operations (much faster)
        /// </summary>
        /// <param name="enabled">True to enable optimization, false to disable</param>
        public void SetOptimization(bool enabled)
        {
            if (corelApp != null)
            {
                try
                {
                    corelApp.Optimization = enabled;
                }
                catch
                {
                    // Ignore optimization errors
                }
            }
        }

        /// <summary>
        /// Disconnect from CorelDRAW and clean up COM resources
        /// This is important to avoid memory leaks
        /// </summary>
        public void Disconnect()
        {
            try
            {
                // Release COM objects
                // This tells Windows we're done using them
                if (activeDocument != null)
                {
                    Marshal.ReleaseComObject(activeDocument);
                    activeDocument = null;
                }

                if (corelApp != null)
                {
                    Marshal.ReleaseComObject(corelApp);
                    corelApp = null;
                }
            }
            catch
            {
                // Ignore disconnect errors
            }
        }
    }
}
