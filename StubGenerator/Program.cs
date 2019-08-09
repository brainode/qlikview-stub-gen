using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.IO;
using System.Text.RegularExpressions;
using CommandLine;

namespace StubGenerator
{
    public class Options
    {
        [Option('p', "path", Required = true, HelpText = "Path to *-prj folder, that contain all project xmls.")]
        public String Path { get; set; }
        [Option('z',"zlevel",Required = true, HelpText = "Set zLevel for Stubs.")]
        public int ZLevel { get; set; }
        [Option('t', "text", Required = true, HelpText = "Set phrase for Stub.")]
        public string Text { get; set; }
        [Option("bgcolor", Required = true, HelpText = "Set background color(you can use rgb code or predefined variable name) for Stubs.")]
        public string BGColor { get; set; }
    }
    class Program
    {
        public static String path = null;
        public const string qvFile = "QlikViewProject.xml";
        
        static Stub createStubFromCH(String chname)
        {
            List<Tuple<String, String>> measures= new List<Tuple<String, String>>();
            var CH = XElement.Load(Path.Combine(path, chname + ".xml"));
            IEnumerable<XElement> expressionsData = CH.Descendants("ArrayOfMainExpressionData");
            foreach(var measure in expressionsData)
            {
                //There is only one definition and expression(in enablecondition) for ArrayOfMainExpressionData element
                var condExpr = measure.Descendants("EnableCondition").First().Element("Expression");
                String expr = Regex.Replace(
                    Regex.Replace(measure.Descendants("Definition").First().Value,
                        "//.*", //Remove all comments
                        String.Empty
                        ),
                    "^=", //Remove first equal symboll
                    String.Empty); 
                String cond = condExpr is null ? null : condExpr.Value;
                measures.Add(new Tuple<String,String>(cond,expr));//TODO. What about named parameters?
            }
            return new Stub(chname, "STUB_" + chname,  measures);
        }
        static void Main(string[] args)
        {
            int zLevel = 0;
            String bgColorVal = "=RGB(0,0,0)";
            String textValue = "No data";

            Parser.Default.ParseArguments<Options>(args)
                   .WithParsed<Options>(o =>
                   {
                       path = o.Path;
                       zLevel = o.ZLevel;
                       bgColorVal = o.BGColor;
                       textValue = o.Text;
                   });
            try
            {
                string qvFilePath = Path.Combine(path, qvFile);
                XElement qvProject = XElement.Load($"{qvFilePath}");
                IEnumerable<XElement> SheetsProperties = qvProject.Descendants("PrjSheetProperties");
                foreach (var sheet in SheetsProperties)
                {
                    List<Stub> stubsInSheet = new List<Stub>();
                    IEnumerable<XElement> sheetElements = sheet.Descendants("PrjFrameParentDef")
                        .Where(sheetItem => Regex.IsMatch(((XElement)sheetItem.FirstNode).Value, "CH") && !Regex.IsMatch(((XElement)sheetItem.FirstNode).Value, "STUB_CH"));
                    foreach (var objInSheet in sheetElements)
                    {
                        //((XElement)objInSheet.FirstNode).Value.Split('\\')[1] - Make split from object ID, to get only name. Default value like Document\\{Needed name}. Is it faster than Regex?
                        Stub stubForCH = createStubFromCH(((XElement)objInSheet.FirstNode).Value.Split('\\')[1]);
                        stubForCH.createStubFile(textValue, bgColorVal);
                        sheet.Element("ChildObjects").Add(new XElement("PrjFrameParentDef",
                                                                    new XElement("ObjectId", @"Document\" + stubForCH.objectID),
                                                                    (XElement)objInSheet.FirstNode.NextNode,//Rect
                                                                    (XElement)objInSheet.FirstNode.NextNode.NextNode,//MinimizedRect
                                                                    new XElement("ShowMode", 1),
                                                                    new XElement("ZedLevel", zLevel)
                                                                    ));
                    }
                }
                //TODO. Put this dogtail in wrapper. Without this bytes at file start Qlikview doesn't work.
                byte[] fileStart = new byte[] { 0xEF, 0xBB, 0xBF, 0x0D, 0x0A };
                using (BinaryWriter writer = new BinaryWriter(File.Open(qvFilePath, FileMode.Create)))
                {
                    writer.Write(fileStart);
                }
                File.AppendAllText(qvFilePath, qvProject.ToString());
            }
            catch (ArgumentNullException)
            {
                Console.WriteLine("There is no QlikViewProject.xml in specified path. The path must contain all *.xml files!");
            }
            
        }
    }
}