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
            AlignSelectedShapes(AngularAlignment.LEFT);
        }

        private void buttonAlignRight_Click(object sender, RibbonControlEventArgs e)
        {
            AlignSelectedShapes(AngularAlignment.RIGHT);
        }

        private void buttonAlignCenter_Click(object sender, RibbonControlEventArgs e)
        {
            AlignSelectedShapes(AngularAlignment.CENTER);
        }


        private void AlignSelectedShapes(AngularAlignment alignment)
        {
            try
            {
                double angle = double.Parse(editBoxAngle.Text) * Math.PI / 180;

                // Get the current selection in PowerPoint
                PPT.Selection selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;

                if (selection.Type == PPT.PpSelectionType.ppSelectionShapes)
                {
                    PPT.ShapeRange selectedShapes = selection.ShapeRange;

                    double minShiftLeft = double.MaxValue;
                    double maxShiftRight = double.MinValue;

                    // Iterate through the selected shapes
                    foreach (PPT.Shape shape in selectedShapes)
                    {

                        // Get the left edge (x) and calculate corresponding y on the line
                        float left = shape.Left;
                        float right = shape.Left + shape.Width;
                        float bottom = shape.Top + shape.Height;
                        float top = shape.Top;

                        double bottomBaseLine = bottom * Math.Tan(-angle);
                        double shiftLeft = left - bottomBaseLine;

                        double topBaseLine = top * Math.Tan(-angle);
                        double shiftRight = right - topBaseLine;

                        minShiftLeft = Math.Min(minShiftLeft, shiftLeft);
                        maxShiftRight = Math.Max(maxShiftRight, shiftRight);
                    }

                    // Iterate through the selected shapes
                    foreach (PPT.Shape shape in selectedShapes)
                    {

                        // Get the left edge (x) and calculate corresponding y on the line
                        float left = shape.Left;
                        float right = shape.Left + shape.Width;
                        float bottom = shape.Top + shape.Height;
                        float top = shape.Top;

                        double targetLine = 0;
                        switch (alignment)
                        {
                            case AngularAlignment.LEFT:
                                targetLine = bottom * Math.Tan(-angle) + minShiftLeft;
                                break;
                            case AngularAlignment.CENTER:
                                targetLine = bottom * Math.Tan(-angle) + minShiftLeft - maxShiftRight;
                                break;
                            case AngularAlignment.RIGHT:
                                targetLine = top * Math.Tan(-angle) + maxShiftRight - shape.Width;
                                break;
                        }
                        

                        // Update shape's Top property to align it along the virtual line
                        shape.Left = (float)targetLine;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Error: {ex.Message}");
            }
        }

        private void buttonStretch_Click(object sender, RibbonControlEventArgs e)
        {

        }
    }

    public enum AngularAlignment
    {
        LEFT,
        RIGHT,
        CENTER
    }
}
