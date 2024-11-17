using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PPT = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PowerPoint.AddIn
{
    public partial class AngularDesignRibbon
    {
        private void AngularDesignRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                double angle = double.Parse(editBoxAngle.Text) * Math.PI / 180;
                // Get the current selection in PowerPoint
                PPT.Selection selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;

                if (selection.Type == PPT.PpSelectionType.ppSelectionShapes)
                {
                    PPT.ShapeRange selectedShapes = selection.ShapeRange;

                    // Iterate through the selected shapes
                    foreach (PPT.Shape shape in selectedShapes)
                    {
                        if (shape.AutoShapeType == Office.MsoAutoShapeType.msoShapeParallelogram)
                        {
                            // Set the slope angle to 77 degrees
                            // The Adjustments[1] range for parallelograms is typically 0 to 1
                            // A value of 0.5 corresponds to no slope. Adjustments vary slightly across versions, test and fine-tune.
                            //double desiredAngleInRadians = -77 * Math.PI / 180; // Convert degrees to radians
                            double desiredAdjustmentValue = Math.Tan(angle); // Approximate slope
                            shape.Adjustments[1] = (float)desiredAdjustmentValue; // Adjustments[1] expects a float

                            //System.Windows.Forms.MessageBox.Show($"Adjusted Parallelogram: {shape.Name}");
                        }
                        else
                        {
                            //System.Windows.Forms.MessageBox.Show($"Skipped non-parallelogram shape: {shape.Name}");
                        }
                    }
                }
                else if (selection.Type == PPT.PpSelectionType.ppSelectionSlides)
                {
                    //System.Windows.Forms.MessageBox.Show("Slides are selected, not shapes.");
                }
                else
                {
                    //System.Windows.Forms.MessageBox.Show("No shapes or slides are selected.");
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Error: {ex.Message}");
            }
        }

        private void buttonAlignLeft_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                double angle = double.Parse(editBoxAngle.Text) * Math.PI / 180;
                
                // Get the current selection in PowerPoint
                PPT.Selection selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;

                if (selection.Type == PPT.PpSelectionType.ppSelectionShapes)
                {
                    PPT.ShapeRange selectedShapes = selection.ShapeRange;

                    double minShift = double.MaxValue;

                    // Iterate through the selected shapes
                    foreach (PPT.Shape shape in selectedShapes)
                    {

                        // Get the left edge (x) and calculate corresponding y on the line
                        float x = shape.Left;
                        float y = shape.Top;
                        double xOnLine = y * Math.Tan(-angle);
                        double shift = shape.Left - xOnLine;
                        if (shift < minShift)
                            minShift = shift;
                    }

                    // Iterate through the selected shapes
                    foreach (PPT.Shape shape in selectedShapes)
                    {

                        // Get the left edge (x) and calculate corresponding y on the line
                        float x = shape.Left;
                        float y = shape.Top;
                        double xOnLine = y * Math.Tan(-angle) + minShift;

                        // Update shape's Top property to align it along the virtual line
                        shape.Left = (float)xOnLine;

                        // Optionally align the shape's left edge to the line
                        // Uncomment the following line if needed:
                        // shape.Left = (float)((shape.Top - yIntercept) / slope);

                    }
                }
                else if (selection.Type == PPT.PpSelectionType.ppSelectionSlides)
                {
                    //System.Windows.Forms.MessageBox.Show("Slides are selected, not shapes.");
                }
                else
                {
                    //System.Windows.Forms.MessageBox.Show("No shapes or slides are selected.");
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Error: {ex.Message}");
            }

        }
    }
}
