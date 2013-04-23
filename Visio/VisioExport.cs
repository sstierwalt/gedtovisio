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
        readonly Random _rand = new Random();
        readonly VisioRenderer _renderer;


        public VisioExport()
        {
            _renderer = new VisioRenderer();
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
            // arrange X on 1x1 grid
            CalcLevel();
            _visualObjects.ForEach(o => o.X = o.Level);


            // arrange Y on 1x1 grid

            var groups = _visualObjects.GroupBy(o => o.Level).OrderBy(g => g.Key).ToList();

            // individuals
            foreach (var levelGroupIndi in groups.Where(o => o.First().GedcomObject is Individual))
            {
                // люди, у которых есть семья, где они родились
                var visualObjectsFromFam = levelGroupIndi.Where(o => o.Parents.Any()).OrderBy(o => o.Parents.Single().Id).ToList();
                // люди, у которых нет информации о семье, где они родились
                var visualObjectsNoFam = levelGroupIndi.Where(o => !o.Parents.Any()).ToList();

                int y = 0;
                foreach (var visualObject in visualObjectsFromFam)
                {
                    visualObject.Y = y++;
                }

                foreach (var visualObject in visualObjectsNoFam)
                {
                    // семья, которую они создали
                    var family = visualObject.Children.FirstOrDefault();
                    if (family == null)
                    {
                        visualObject.Y = y++;
                        continue;
                    }
                    // супруг
                    var spouse = family.Parents.FirstOrDefault(o => o != visualObject);
                    if (spouse == null)
                    {
                        visualObject.Y = y++;
                        continue;
                    }
                    visualObject.Y = spouse.Y + 1;
                    foreach (var o in levelGroupIndi.Where(o => o.Y >= visualObject.Y && o != visualObject))
                    {
                        o.Y++;
                    }
                    y++;
                }
            }


            // family
            foreach (var levelGroupFam in groups.Where(o => o.First().GedcomObject is Family))
            {
                foreach (var visualObject in levelGroupFam)
                {
                    if (visualObject.Children.Any())
                    {
                        visualObject.Y = (int)visualObject.Children.Average(o => o.Y);
                    }
                }
            }

            //var visualObjects = _visualObjects.OrderBy(o => o.Level).ToList();


            //foreach (var visualObject in visualObjects)
            //{
            //    var childCount = visualObject.Children.Count;
            //    for (int i = 0; i < childCount; i++)
            //    {
            //        visualObject.Children[i].Y = visualObject.Y + i;
            //    }
            //}

            //foreach (var visualObject in visualObjects.OrderByDescending(o => o.Level))
            //{
            //    var husb = visualObject.Husband;
            //    var wife = visualObject.Wife;
            //    if (husb == null || wife == null)
            //    {
            //        continue;
            //    }

            //    if (husb.Y == wife.Y)
            //    {
            //        husb.Y = visualObject.Y + 1;
            //        wife.Y = visualObject.Y - 1;
            //    }
            //}

            Render();

            //var cost = Cost();
            //for (int i = 0; i < 10000; i++)
            //{
            //    // random objects to move
            //    var objects = new List<VisualObject>();
            //    var n = 2;
            //    for (int j = 0; j < n; j++)
            //    {
            //        // random object
            //        var o = _visualObjects[_rand.Next(_visualObjects.Count)];
            //        // save position
            //        o.OrigY = o.Y;
            //        // random move
            //        o.Y += _rand.Next(11) - 5;

            //        objects.Add(o);
            //    }

            //    var newCost = Cost();
            //    if (cost > newCost)
            //    {
            //        // apply
            //        cost = newCost;
            //        objects.ForEach(o => _renderer.Move(o.Shape, o.X, o.Y));
            //    }
            //    else
            //    {
            //        // undo
            //        objects.ForEach(o => o.Y = o.OrigY);
            //    }
            //}
        }


        double Cost()
        {
            // штраф за совпадающие позиции
            int overlappingCount = _visualObjects.Sum(visualObject => _visualObjects.Count(o => o.X == visualObject.X && o.Y == visualObject.Y && o != visualObject))/2;

            // длинна связей
            double len = _visualObjects.Sum(o => o.Children.Sum(c => Math.Sqrt((c.X - o.X)*(c.X - o.X) + (c.Y - o.Y)*(c.Y - o.Y))));
            return overlappingCount * 10 + len;
        }


        /// <summary>
        /// Определяет уровни дерева, соответствующие поколениям (и семьям).
        /// </summary>
        public void CalcLevel()
        {
            bool hasChanges;

            do
            {
                hasChanges = false;
                // "раздвигаем" перекрывающиеся уровни
                foreach (var o in _visualObjects)
                {
                    var clildLevel = o.Level + 1;
                    foreach (var child in o.Children.Where(child => child.Level < clildLevel))
                    {
                        child.Level = clildLevel;
                        hasChanges = true;
                    }
                }
                // "придвигаем" объекты, которые отстоят слишком далеко от детей
                foreach (var o in _visualObjects)
                {
                    if (o.Children.Any())
                    {
                        var level = o.Children.Max(c => c.Level) - 1;
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
                var shape = _renderer.Render(visualObject.GedcomObject, visualObject.X, visualObject.Y);
                visualObject.Shape = shape;               
            }

            foreach (var visualObject in _visualObjects.Where(o => o.GedcomObject is Family))
            {
                var shape = _renderer.Render(visualObject.GedcomObject, visualObject.X, visualObject.Y);
                visualObject.Shape = shape;
                foreach (var parent in visualObject.Parents)
                {
                    _renderer.Connect(shape, parent.Shape);
                }

                foreach (var children in visualObject.Children)
                {
                    _renderer.Connect(shape, children.Shape);
                }
            }
        }

    }
}
