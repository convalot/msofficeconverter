using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace mspowerpointconverter
{
    class Program
    {
        static void Main(string[] args)
        {
            if(args.Length < 2 || args.Length > 3)
            {
                System.Console.WriteLine("Error: no filenames specified");
                System.Console.WriteLine("Usage: mspowerpointconverter inputfile outputfile <saveformat>");
                System.Console.WriteLine("saveformat is optional, will try to autodetect format and fall back to default if not given.");
                System.Console.WriteLine("supports output to ppt, pptx, html, odp, pdf, xps, xml, wmv");
                System.Console.WriteLine("will also export an image of each slide in bmp, jpg, png, tif, emf, or gif");
                System.Console.WriteLine("and can export an outline as rtf.");
                System.Console.WriteLine("Support for a given format depends on whether support is present in office itself, notably for pdf.");
                return;
            }
            //read filenames
            StringBuilder input = new StringBuilder(args[0].Length);
            for (int i = 0; i < args[0].Length; i++)
            {
                if (args[0][i] == '/')
                    input.Append('\\');
                else
                    input.Append(args[0][i]);
            }
            StringBuilder output = new StringBuilder(args[1].Length);
            for (int i = 0; i < args[1].Length; i++)
            {
                if (args[1][i] == '/')
                    output.Append('\\');
                else
                    output.Append(args[1][i]);
            }
            //select output format
            PpSaveAsFileType format = PpSaveAsFileType.ppSaveAsDefault;
            string formatString;
            if (args.Length == 3)
            {
                formatString = args[2];
            }
            else
            {
                formatString = output.ToString().Substring(output.ToString().LastIndexOf(".") + 1);
            }
            if (formatString == "ppt")
                format = PpSaveAsFileType.ppSaveAsPresentation;
            if (formatString == "pptx")
                format = PpSaveAsFileType.ppSaveAsOpenXMLPresentation;
            if (formatString == "rtf")
                format = PpSaveAsFileType.ppSaveAsRTF;
            // Powerpoint has three different HTML output formats. The uncommented one seemed to be the least dependent on ActiveX
            if (formatString == "html")
                format = PpSaveAsFileType.ppSaveAsHTML;
           // if (formatString == "html")
             //   format = PpSaveAsFileType.ppSaveAsHTMLv3;
           // if (formatString == "html")
             //   format = PpSaveAsFileType.ppSaveAsHTMLDual;
            if (formatString == "pps")
                format = PpSaveAsFileType.ppSaveAsShow;
            if (formatString == "ppsx")
                format = PpSaveAsFileType.ppSaveAsOpenXMLShow;
            if (formatString == "pdf")
                format = PpSaveAsFileType.ppSaveAsPDF;
            if (formatString == "xps")
                format = PpSaveAsFileType.ppSaveAsXPS;
            if (formatString == "bmp")
                format = PpSaveAsFileType.ppSaveAsBMP;
            if (formatString == "jpg")
                format = PpSaveAsFileType.ppSaveAsJPG;
            if (formatString == "png")
                format = PpSaveAsFileType.ppSaveAsPNG;
            if (formatString == "gif")
                format = PpSaveAsFileType.ppSaveAsGIF;
            if (formatString == "emf")
                format = PpSaveAsFileType.ppSaveAsEMF;
            if (formatString == "tif")
                format = PpSaveAsFileType.ppSaveAsTIF;
            if (formatString == "wmv")
                format = PpSaveAsFileType.ppSaveAsWMV;
            if (formatString == "xml")
                format = PpSaveAsFileType.ppSaveAsXMLPresentation;
            
            
            Application app;
            try
            {
                app = new Application();
            }
            catch(Exception ex)
            {
                System.Console.WriteLine("Unable to open Microsoft Word");
                System.Console.WriteLine("Error: " + ex.Message);
                return;
            }
            Presentation pres;
            try
            {
                pres = app.Presentations.Open(input.ToString(), MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);
            }
            catch(Exception ex)
            {
                System.Console.WriteLine("Unable to open file " + input);
                System.Console.WriteLine("Error: " + ex.Message);
                app.Quit();
                return;
            }
            if (pres != null)
            {
                try
                {
                    //doc.SaveAs2(output, format);
                    pres.SaveAs(output.ToString(), format, MsoTriState.msoTriStateMixed);
                }
                catch (Exception ex)
                {
                    System.Console.WriteLine("Save to " + output + " failed");
                    System.Console.WriteLine("Error: " + ex.Message);
                }
            }
            else
            {
                System.Console.WriteLine("unable to open file");
            }

            pres.Close();
            app.Quit();
        }
    }
}
