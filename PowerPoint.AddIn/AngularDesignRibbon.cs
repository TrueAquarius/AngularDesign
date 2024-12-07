using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PPT = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPoint.AddIn
{
    public partial class AngularDesignRibbon
    {
        private void AngularDesignRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }



        /// <summary>
        /// Event handler for button "Apply"
        /// </summary>
        /// <param name="sender">Sender</param>
        /// <param name="e">Event Arguents</param>
        private void buttonApply_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // Read the angle from input box
                float angle = (float)(double.Parse(editBoxAngle.Text) * Math.PI / 180);

                // Apply angle to selected shapes 
                ApplyAngle(Globals.ThisAddIn.Application.ActiveWindow.Selection, angle);

            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"An error has occured in Plug-In 'Angular Design': {ex.Message}");
            }
        }



        /// <summary>
        /// Event handler for button "Pick angle"
        /// </summary>
        /// <param name="sender">Sender</param>
        /// <param name="e">Event Arguents</param>
        private void buttonPick_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // Get the current selection in PowerPoint
                PPT.Selection selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;

                if (selection != null && selection.Type == PPT.PpSelectionType.ppSelectionShapes)
                {
                    PPT.ShapeRange selectedShapes = selection.ShapeRange;

                    if (selectedShapes.Count > 0)
                    {
                        Shape shape = selectedShapes[1];

                        float angle = GetAngle(shape);

                        editBoxAngle.Text = Math.Round(angle * 180 / Math.PI, 1).ToString();

                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"An error has occured in Plug-In 'Angular Design': {ex.Message}");
            }
        }


        /// <summary>
        /// Event handler for button "Align Left"
        /// </summary>
        /// <param name="sender">Sender</param>
        /// <param name="e">Event Arguents</param>
        private void buttonAlignLeft_Click(object sender, RibbonControlEventArgs e)
        {
            AlignSelectedShapes(AngularAlignment.LEFT);
        }


        /// <summary>
        /// Event handler for button "Align Right"
        /// </summary>
        /// <param name="sender">Sender</param>
        /// <param name="e">Event Arguents</param>
        private void buttonAlignRight_Click(object sender, RibbonControlEventArgs e)
        {
            AlignSelectedShapes(AngularAlignment.RIGHT);
        }

        /// <summary>
        /// Event handler for button "Align Center"
        /// </summary>
        /// <param name="sender">Sender</param>
        /// <param name="e">Event Arguents</param>
        private void buttonAlignCenter_Click(object sender, RibbonControlEventArgs e)
        {
            AlignSelectedShapes(AngularAlignment.CENTER);
        }

        /// <summary>
        /// Event handler for button "Align Stretch"
        /// </summary>
        /// <param name="sender">Sender</param>
        /// <param name="e">Event Arguents</param>
        private void buttonStretch_Click(object sender, RibbonControlEventArgs e)
        {
            AlignSelectedShapes(AngularAlignment.STRETCH);
        }



 


        /// <summary>
        /// Aligns selected shapes as specified.
        /// </summary>
        /// <param name="alignment">Target alignmant</param>
        private void AlignSelectedShapes(AngularAlignment alignment)
        {
            try
            {
                // Get the current selection in PowerPoint
                PPT.Selection selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;

                // Read the angle from input box
                float angle = (float)(double.Parse(editBoxAngle.Text) * Math.PI / 180);

                // Apply angle to selected shapes 
                ApplyAngle(selection, angle);

                if (selection.Type == PPT.PpSelectionType.ppSelectionShapes)
                {
                    PPT.ShapeRange selectedShapes = selection.ShapeRange;

                    double minShiftLeft = double.MaxValue;
                    double maxShiftRight = double.MinValue;
                    double sumShiftCenter = 0;
                    float maxWidth = 0;

                    int i = 0;
                    // Iterate through the selected shapes in order to determine maxima and minima
                    foreach (PPT.Shape shape in selectedShapes)
                    {
                        ++i;

                        // Get the left edge (x) and calculate corresponding y on the line
                        float left = shape.Left;
                        float right = shape.Left + shape.Width;
                        float bottom = shape.Top + shape.Height;
                        float top = shape.Top;
                        float center = (top + bottom) / 2;
                        float middle = (left + right) / 2;

                        double bottomBaseLine = bottom * Math.Tan(-angle);
                        double shiftLeft = left - bottomBaseLine;

                        double topBaseLine = top * Math.Tan(-angle);
                        double shiftRight = right - topBaseLine;

                        double centerBaseLine = center * Math.Tan(-angle);
                        double shiftCenter = middle - centerBaseLine;

                        minShiftLeft = Math.Min(minShiftLeft, shiftLeft);
                        maxShiftRight = Math.Max(maxShiftRight, shiftRight);
                        sumShiftCenter += shiftCenter;

                        float width =shape.Width - shape.Height * (float)Math.Tan(angle);
                        maxWidth = Math.Max(maxWidth, width);
                    }

                    sumShiftCenter /= i;

                    // Iterate through the selected shapes in order shift left or right or resize
                    foreach (PPT.Shape shape in selectedShapes)
                    {
                        // Get the left edge (x) and calculate corresponding y on the line
                        float left = shape.Left;
                        float right = shape.Left + shape.Width;
                        float bottom = shape.Top + shape.Height;
                        float top = shape.Top;
                        float center = (top + bottom) / 2;
                        float middle = (left + right) / 2;
                        float width = shape.Width - shape.Height * (float)Math.Tan(angle);
                        double targetLine = 0;
                        switch (alignment)
                        {
                            case AngularAlignment.LEFT:
                                targetLine = bottom * Math.Tan(-angle) + minShiftLeft;
                                break;
                            case AngularAlignment.RIGHT:
                                targetLine = top * Math.Tan(-angle) + maxShiftRight - shape.Width;
                                break;
                            case AngularAlignment.CENTER:
                                targetLine = center * Math.Tan(-angle) + sumShiftCenter - shape.Width / 2;
                                break;
                            case AngularAlignment.STRETCH:
                                shape.Width = maxWidth + shape.Height * (float)Math.Tan(angle);
                                targetLine = center * Math.Tan(-angle) + sumShiftCenter - shape.Width / 2;
                                break;
                        }

                        // Update shape's Left property to align it along the virtual line
                        shape.Left = (float)targetLine;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"An error has occured in Plug-In 'Angular Design': {ex.Message}");
            }
        }



        private void ApplyAngle(PPT.Selection selection, float angle)
        {
            try
            {
                if (selection.Type == PPT.PpSelectionType.ppSelectionShapes)
                {
                    PPT.ShapeRange selectedShapes = selection.ShapeRange;

                    // Iterate through the selected shapes
                    foreach (PPT.Shape shape in selectedShapes)
                    {
                        ApplyAngle(shape, angle);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"An error has occured in Plug-In 'Angular Design': {ex.Message}");
            }
        }



        private void ApplyAngle(Shape shape, float angle)
        {  
            switch (shape.AutoShapeType)
            {
                case Office.MsoAutoShapeType.msoShapeParallelogram:
                    shape.Adjustments[1] = angle;
                    break;
                default:      
                    break;
            }
        }




        private float GetAngle(Shape shape)
        {
            float angle = 0;

            switch (shape.AutoShapeType)
            {
                case Office.MsoAutoShapeType.msoShapeParallelogram:
                    angle = shape.Adjustments[1]; 
                    break;
                default:
                    angle = 0;
                    break;
            }

            return angle;
        }

    }




    public enum AngularAlignment
    {
        LEFT,
        RIGHT,
        CENTER,
        STRETCH
    }
}
