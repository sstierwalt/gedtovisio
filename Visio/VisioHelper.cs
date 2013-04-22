using System;
using Microsoft.Office.Interop.Visio;

namespace GedToVisio.Visio
{
    public static class VisioHelper
    {
        public static void ConnectWithDynamicGlueAndConnector(
                    Microsoft.Office.Interop.Visio.Shape shapeFrom,
                    Microsoft.Office.Interop.Visio.Shape shapeTo)
        {


            if (shapeFrom == null || shapeTo == null)
            {
                return;
            }


            const string BASIC_FLOWCHART_STENCIL =
                "Basic Flowchart Shapes (US units).vss";
            const string DYNAMIC_CONNECTOR_MASTER = "Dynamic Connector";
            const string MESSAGE_NOT_SAME_PAGE =
                "Both the shapes are not on the same page.";

            Microsoft.Office.Interop.Visio.Application visioApplication;
            Microsoft.Office.Interop.Visio.Document stencil;
            Microsoft.Office.Interop.Visio.Master masterInStencil;
            Microsoft.Office.Interop.Visio.Shape connector;
            Microsoft.Office.Interop.Visio.Cell beginX;
            Microsoft.Office.Interop.Visio.Cell endX;

            // Get the Application object from the shape.
            visioApplication = (Microsoft.Office.Interop.Visio.Application)
                shapeFrom.Application;

            try
            {

                // Verify that the shapes are on the same page.
                if (shapeFrom.ContainingPage != null && shapeTo.ContainingPage != null &&
                    shapeFrom.ContainingPage.Equals(shapeTo.ContainingPage))
                {

                    // Access the Basic Flowchart Shapes stencil from the
                    // Documents collection of the application.
                    stencil = visioApplication.Documents.OpenEx(
                        BASIC_FLOWCHART_STENCIL,
                        (short)Microsoft.Office.Interop.Visio.
                            VisOpenSaveArgs.visOpenDocked);

                    // Get the dynamic connector master on the stencil by its
                    // universal name.
                    masterInStencil = stencil.Masters.get_ItemU(
                        DYNAMIC_CONNECTOR_MASTER);

                    // Drop the dynamic connector on the active page.
                    connector = visioApplication.ActivePage.Drop(
                        masterInStencil, 0, 0);

                    // Connect the begin point of the dynamic connector to the
                    // PinX cell of the first 2-D shape.
                    beginX = connector.get_CellsSRC(
                        (short)Microsoft.Office.Interop.Visio.
                            VisSectionIndices.visSectionObject,
                        (short)Microsoft.Office.Interop.Visio.
                            VisRowIndices.visRowXForm1D,
                        (short)Microsoft.Office.Interop.Visio.
                            VisCellIndices.vis1DBeginX);

                    beginX.GlueTo(shapeFrom.get_CellsSRC(
                        (short)Microsoft.Office.Interop.Visio.
                            VisSectionIndices.visSectionObject,
                        (short)Microsoft.Office.Interop.Visio.
                            VisRowIndices.visRowXFormOut,
                        (short)Microsoft.Office.Interop.Visio.
                            VisCellIndices.visXFormPinX));

                    // Connect the end point of the dynamic connector to the
                    // PinX cell of the second 2-D shape.
                    endX = connector.get_CellsSRC(
                        (short)Microsoft.Office.Interop.Visio.
                            VisSectionIndices.visSectionObject,
                        (short)Microsoft.Office.Interop.Visio.
                            VisRowIndices.visRowXForm1D,
                        (short)Microsoft.Office.Interop.Visio.
                            VisCellIndices.vis1DEndX);

                    endX.GlueTo(shapeTo.get_CellsSRC(
                        (short)Microsoft.Office.Interop.Visio.VisSectionIndices.visSectionObject,
                        (short)Microsoft.Office.Interop.Visio.VisRowIndices.visRowXFormOut,
                        (short)Microsoft.Office.Interop.Visio.VisCellIndices.visXFormPinX));

                    connector.CellsU["ShapeRouteStyle"].FormulaU = "16";
//                    var x = (short)Microsoft.Office.Interop.Visio.VisCellVals.visLORouteCenterToCenter;
                }
                else
                {

                    // Processing cannot continue because the shapes are not on 
                    // the same page.
                    System.Diagnostics.Debug.WriteLine(MESSAGE_NOT_SAME_PAGE);
                }
            }
            catch (Exception err)
            {
                System.Diagnostics.Debug.WriteLine(err.Message);
                throw;
            }
        }


        public static string GetCustomProperty(Shape shape, string propertyName)
        {
            var propName = "Prop." + propertyName;

            var customPropertyCell = shape.CellsU[propName];
            return customPropertyCell.ResultStr[VisUnitCodes.visNoCast];
        }


        public static void SetCustomProperty(Shape shape, string propertyName, string value) 
        {
            SetCustomProperty(shape, propertyName, propertyName, value);
        }


        public static void SetCustomProperty(Shape shape, string propertyName, string label, string value)
        {
            short rowIndex = shape.AddNamedRow((short)VisSectionIndices.visSectionProp, propertyName,
                                               (short)VisRowIndices.visRowProp);

            shape.CellsSRC[(short)VisSectionIndices.visSectionProp, rowIndex, (short)VisCellIndices.visCustPropsLabel].
                FormulaU = StringToFormulaForString(label);

            shape.CellsSRC[(short)VisSectionIndices.visSectionProp, rowIndex, (short)VisCellIndices.visCustPropsType].FormulaU
                =
                StringToFormulaForString(
                    ((short)VisCellVals.visPropTypeString).ToString(System.Globalization.CultureInfo.InvariantCulture));

            shape.CellsSRC[(short)VisSectionIndices.visSectionProp, rowIndex, (short)VisCellIndices.visCustPropsValue].
                FormulaU = StringToFormulaForString(value);
        }

        public static string StringToFormulaForString(string inputValue)
        {
            const string quote = "\"";
            const string quoteQuote = "\"\"";

            return quote + (inputValue ?? String.Empty).Replace(quote, quoteQuote) + quote;
        }

    }
}
