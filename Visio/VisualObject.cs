﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using GedToVisio.Gedcom;
using Microsoft.Office.Interop.Visio;

namespace GedToVisio.Visio
{
    public class VisualObject
    {
        public string Id { get; set; }

        /// <summary>
        /// Уровень в дереве, начиная от корня
        /// </summary>
        public int Level { get; set; }

        public int X { get; set; }
        public int Y { get; set; }
        public int OrigY { get; set; }
        public Shape Shape { get; set; }
        public Gedcom.GedcomObject GedcomObject { get; set; }

        public List<VisualObject> Parents { get; set; }

        public List<VisualObject> Children { get; set; }

        /// <summary>
        /// Муж, если объект - семья.
        /// </summary>
        public VisualObject Husband { get; set; }

        /// <summary>
        /// Жена, если объект - семья.
        /// </summary>
        public VisualObject Wife { get; set; }


        string IndiStr(Individual indi)
        {
            return string.Format("{0} {1}", indi.GivenName, indi.Surname);
        }

        public override string ToString()
        {
            string s;
            var indi = GedcomObject as Individual;
            if (indi != null)
            {
                s = IndiStr(indi);
            }
            else
            {
                s = string.Join(" - ", Parents.Select(p => (Individual)p.GedcomObject).Select(IndiStr));
            }
            return string.Format("[{0}] {1} {2} {3}", Level, X, Y, s);
        }
    }
}
