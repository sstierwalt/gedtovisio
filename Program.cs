using System.Text;
using GedToVisio.Gedcom;
using GedToVisio.Visio;

namespace GedToVisio
{
    class Program
    {
        static void Main(string[] args)
        {
            var gedcomFile = new GedcomFile();
            var path = System.IO.Path.GetFullPath(args[0]);
            gedcomFile.Load(path, Encoding.GetEncoding(1251));

            var visioExport = new VisioExport();

            gedcomFile.Individuals.ForEach(visioExport.Add);
            gedcomFile.Families.ForEach(visioExport.Add);

            visioExport.BuildTree();

            visioExport.Arrange();

            //visioExport.Render();
        }
    }
}
