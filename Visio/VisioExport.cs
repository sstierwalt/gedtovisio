using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using GedToVisio.Gedcom;
using Microsoft.Office.Interop.Visio;

namespace GedToVisio.Visio
{
    public class VisioExport
    {
        private readonly List<VisualObject> _visualObjects = new List<VisualObject>();
        readonly Page _page;
        readonly Random _rand = new Random();


        public VisioExport()
        {
            var application = new Microsoft.Office.Interop.Visio.Application();
            application.Documents.Add("");
            _page = application.Documents[1].Pages[1];
        }


        public void Add(GedcomObject indi)
        {
            var visualObject = new VisualObject { Id = indi.Id, GedcomObject = indi, Children = new List<VisualObject>(), Parents = new List<VisualObject>() };
            _visualObjects.Add(visualObject);
        }


        /// <summary>
        /// Строит дерево загруженых объектов, заполняя Parents, Children, Husband, Wife.
        /// </summary>
        public void BuildTree()
        {
            var visualObjectsDict = _visualObjects.ToDictionary(o => o.GedcomObject.Id);

            // обходим все семьи
            foreach (var visualObject in _visualObjects)
            {
                var fam = visualObject.GedcomObject as Family;
                if (fam == null)
                    continue;

                if (fam.HusbandId != null && visualObjectsDict.ContainsKey(fam.HusbandId))
                {
                    var husband = visualObjectsDict[fam.HusbandId];
                    visualObject.Husband = husband;
                    visualObject.Parents.Add(husband);
                    husband.Children.Add(visualObject);
                }

                if (fam.WifeId != null && visualObjectsDict.ContainsKey(fam.WifeId))
                {
                    var wife = visualObjectsDict[fam.WifeId];
                    visualObject.Wife = wife;
                    visualObject.Parents.Add(wife);
                    wife.Children.Add(visualObject);
                }

                foreach (var childId in fam.Children)
                {
                    if (visualObjectsDict.ContainsKey(childId))
                    {
                        var child = visualObjectsDict[childId];
                        visualObject.Children.Add(child);
                        child.Parents.Add(visualObject);
                    }
                }
            }
        }


        /// <summary>
        /// Расположить объекты на листе.
        /// </summary>
        public void Arrange()
        {
            // arrange X
            CalcLevel();

            foreach (var visualObject in _visualObjects)
            {
                visualObject.X = 10 - visualObject.Level * 2;
            }


            // arrange Y
            var visualObjects = _visualObjects.OrderBy(o => o.Level).ToList();

            foreach (var visualObject in _visualObjects)
            {
                visualObject.Y = _rand.NextDouble() * 11;
            }

            foreach (var visualObject in _visualObjects)
            {
                if (visualObject.Husband != null)
                {
                    visualObject.Husband.Y = visualObject.Y + 1;
                }

                if (visualObject.Wife != null)
                {
                    visualObject.Wife.Y = visualObject.Y - 1;
                }
            }
        }



        public void CalcLevel()
        {
            bool hasChanges;

            // "раздвигаем" перекрывающиеся уровни
            do
            {
                hasChanges = false;
                foreach (var o in _visualObjects)
                {
                    var clildLevel = o.Level + 1;
                    foreach (var child in o.Children.Where(child => child.Level < clildLevel))
                    {
                        child.Level = clildLevel;
                        hasChanges = true;
                    }
                }
                
            } while (hasChanges);

            // "придвигаем" объекты, которые отстоят слишком далеко от детей
            do
            {
                hasChanges = false;
                foreach (var o in _visualObjects)
                {
                    if (o.Children.Any())
                    {
                        var level = o.Children.Min(c => c.Level) - 1;
                        if (level > o.Level)
                        {
                            o.Level = level;
                            hasChanges = true;
                        }
                    }
                }
            } while (hasChanges);
        }


        public void Render()
        {
            foreach (var visualObject in _visualObjects.Where(o => o.GedcomObject is Individual))
            {
                var shape = Render(visualObject.GedcomObject, visualObject.X, visualObject.Y);
                visualObject.Shape = shape;               
            }

            foreach (var visualObject in _visualObjects.Where(o => o.GedcomObject is Family))
            {
                var shape = Render(visualObject.GedcomObject, visualObject.X, visualObject.Y);
                visualObject.Shape = shape;
                foreach (var parent in visualObject.Parents)
                {
                    VisioHelper.ConnectWithDynamicGlueAndConnector(shape, parent.Shape);
                }

                foreach (var children in visualObject.Children)
                {
                    VisioHelper.ConnectWithDynamicGlueAndConnector(shape, children.Shape);
                }

            }
        }

    
        Shape Render(object o, double x, double y)
        {
            if (o is Individual)
            {
                return Render((Individual) o, x, y);
            }

            if (o is Family)
            {
                return Render((Family) o, x, y);
            }
            throw new ArgumentException();
        }


        Shape Render(Individual indi, double x, double y)
        {
            var height = 0.5;
            var width = 1.7;

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
            var height = 0.4;
            var width = 0.8;

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
