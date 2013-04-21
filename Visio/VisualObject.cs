using System;
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
        public int Level { get; set; }
        public double X { get; set; }
        public double Y { get; set; }
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


        public override string ToString()
        {
            string s;
            var indi = GedcomObject as Individual;
            if (indi != null)
            {
                s = string.Format("{0} {1}", indi.GivenName, indi.Surname);
            }
            else
            {
                s = string.Join(" - ", Parents.Select(p => (Individual) p.GedcomObject).Select(p => string.Format("{0} {1}", p.GivenName, p.Surname)));
            }
            return string.Format("{0} {1}", Level, s);
        }
    }
}
