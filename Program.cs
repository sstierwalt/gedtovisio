﻿using System;
using System.Text;
using GedToVisio.Gedcom;
using GedToVisio.Visio;

namespace GedToVisio
{
    class Program
    {
        static void Main(string[] args)
        {
            // parse command line
            string inputGedcomFilePath;
            int encoding;
            GedcomFile gedcomFile;
            VisioExport visioExport;

            if (args.Length < 1)
            {
                Console.WriteLine("GedToVisio <inputGedcomFile.ged> [encoding]");
                return;
            }
            inputGedcomFilePath = System.IO.Path.GetFullPath(args[0]);

            try
            {
                encoding = args.Length < 2 ? 1251 : int.Parse(args[1]);
            }
            catch (Exception)
            {
                Console.WriteLine("Incorrect encoding - number expected, for example 1251");
                return;
            }

            // parse GEDCOM
            try
            {
                Console.WriteLine("Parsing GEDCOM file: {0}", inputGedcomFilePath);
                gedcomFile = new GedcomFile();
                gedcomFile.Load(inputGedcomFilePath, Encoding.GetEncoding(encoding));
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error parsing GEDCOM file: {0}", ex);
                return;
            }

            // open Visio
            Console.WriteLine("Exporting to MS Visio");
            try
            {
                visioExport = new VisioExport();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Can't open MS Visio: {0}", ex);
                return;
            }

            // export
            try
            {
                gedcomFile.Individuals.ForEach(visioExport.Add);
                gedcomFile.Families.ForEach(visioExport.Add);

                visioExport.BuildTree();

                visioExport.Arrange();

                //visioExport.Render();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error exporting to Visio: {0}", ex);
                return;
            }

            Console.WriteLine("Done.");
        }
    }
}
