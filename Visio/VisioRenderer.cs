using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using GedToVisio.Gedcom;
using Microsoft.Office.Interop.Visio;

namespace GedToVisio.Visio
{
    /// <summary>
    /// Отображает генеалогические объекты в документе MS Visio.
    /// </summary>
    public class VisioRenderer
    {
        readonly Page _page;

        private static double ScaleX(int x)
        {
            return 10.0 - x * 2;
        }

        private static double ScaleY(int y)
        {
            return 6.0 + y;
        }


        public VisioRenderer()
        {
            var application = new Microsoft.Office.Interop.Visio.Application();
            application.Documents.Add("");
            _page = application.Documents[1].Pages[1];            
        }


        public Shape Render(GedcomObject o, int x, int y)
        {
            if (o is Individual)
            {
                return Render((Individual)o, ScaleX(x), ScaleY(y));
            }

            if (o is Family)
            {
                return Render((Family)o, ScaleX(x), ScaleY(y));
            }
            throw new ArgumentException();
        }


        public void Connect(Shape shape1, Shape shape2)
        {
            VisioHelper.ConnectWithDynamicGlueAndConnector(shape1, shape2);
        }


        public void Move(Shape shape, int x, int y)
        {
            shape.SetCenter(ScaleX(x), ScaleY(y));           
        }


        Shape Render(Individual indi, double x, double y)
        {
            const double height = 0.5;
            const double width = 1.7;

            var shape = _page.DrawRectangle(x - width / 2, y - height / 2, x + width / 2, y + height / 2);
            shape.Text = string.Format("{0} /{1}/\n{2} - {3}", indi.GivenName, indi.Surname, indi.BirthDate, indi.DiedDate);

            shape.Characters.CharProps[(short)VisCellIndices.visCharacterSize] = 10;
            if (indi.Sex == "F")
            {
                shape.CellsU["Rounding"].FormulaU = ".2";
            }

            VisioHelper.SetCustomProperty(shape, "_UID", "Unique Identification Number", indi.Uid);
            VisioHelper.SetCustomProperty(shape, "ID", "gedcom ID", indi.Id);
            VisioHelper.SetCustomProperty(shape, "GivenName", indi.GivenName);
            VisioHelper.SetCustomProperty(shape, "Surname", indi.Surname);
            VisioHelper.SetCustomProperty(shape, "Sex", indi.Sex);
            VisioHelper.SetCustomProperty(shape, "BirthDate", indi.BirthDate);
            VisioHelper.SetCustomProperty(shape, "DiedDate", indi.DiedDate);

            return shape;
        }


        Shape Render(Family fam, double x, double y)
        {
            const double height = 0.4;
            const double width = 0.8;

            var shape = _page.DrawOval(x - width / 2, y - height / 2, x + width / 2, y + height / 2);
            shape.Text = string.Format("{0}", fam.MarriageDate);
            shape.Characters.CharProps[(short)VisCellIndices.visCharacterSize] = 8;

            VisioHelper.SetCustomProperty(shape, "_UID", "Unique Identification Number", fam.Uid);
            VisioHelper.SetCustomProperty(shape, "ID", "gedcom ID", fam.Id);
            VisioHelper.SetCustomProperty(shape, "MarriageDate", fam.MarriageDate);

            return shape;
        }    
    }
}
