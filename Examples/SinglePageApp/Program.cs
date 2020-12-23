﻿using System;
using Aspose.Slides;
using Aspose.Slides.Export.Web;
using Aspose.Slides.WebExtensions;

namespace SinglePageApp
{
    class Program
    {
        static void Main(string[] args)
        {
            using (Presentation pres = new Presentation("demo-transitions.pptx"))
            {
                WebDocument document = pres.ToSinglePageWebDocument("templates\\single-page",
                    @"single-page-output");
                document.Save();
            }

            Console.WriteLine("HTML export is complete...");
        }
    }
}
