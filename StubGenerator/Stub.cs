using System;
using System.Collections.Generic;
using System.Xml.Linq;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StubGenerator
{
    class Stub
    {
        String sheetID;
        public String objectID;
        List<Tuple<String,String>> measures;
        String getCondition()
        {
            StringBuilder answer = new StringBuilder();
            answer.Append("if(");
            bool prevMeasureWCond = false;
            foreach(var measure in measures)
            {
                if (measure != measures.First())
                {
                    if (prevMeasureWCond)
                    {
                        answer.Append(" or ");
                    }
                    else
                    {
                        answer.Append(" and ");
                    }
                }
                //TODO. What about named parameters?Check Program.cs
                if (measure.Item1 != null)
                {
                    prevMeasureWCond = true;
                    answer.AppendFormat("({0} and alt({1})=0)",measure.Item1,measure.Item2);
                }
                else
                {
                    answer.AppendFormat("(alt({0},0)=0)", measure.Item2);
                }
            }
            answer.Append(",1,0)");
            return answer.ToString();
        }
        public Stub(string sheetID,string objectID, List<Tuple<String,String>> measures)
        {
            this.sheetID = sheetID;
            this.objectID = objectID;
            this.measures = measures;
        }
        public void createStubFile(String textVal,String bgColor,String textColor=null)
        {
            //TODO. Here must be code, that inserts information about TX object in QlikViewProject.xml
            var stubFile = new XElement("TextObjectProperties",
                                new XElement("Layout",
                                    new XElement("Frame",
                                        new XElement("BorderWidth", "0"),
                                        new XElement("BorderEffect", "FLAT"),
                                        new XElement("Color",
                                            new XElement("Mode", "COLORAREA_SOLID"),
                                            new XElement("Luminosity", "85"),
                                            new XElement("FillDirection", "COLORAREA_HORZ"),
                                            new XElement("FillPattern", "COLORAREA_PATTERN1"),
                                            new XElement("PrimaryCol",
                                                new XElement("Col",
                                                    new XElement("Red", "0"),
                                                    new XElement("Green", "98"),
                                                    new XElement("Blue", "164"),
                                                    new XElement("Alpha", "255")
                                                ),
                                                new XElement("IsCalculated", "false"),
                                                new XElement("ColorExpr", "")
                                            ),
                                            new XElement("SecondaryCol", ""),
                                            new XElement("Alpha", "255")
                                        ),
                                        new XElement("Font", ""),
                                        new XElement("CaptionFont",
                                            new XElement("FontName", "Arial"),
                                            new XElement("PointSize1000", "10000"),
                                            new XElement("Underline", "false"),
                                            new XElement("Bold", "true"),
                                            new XElement("Italic", "false"),
                                            new XElement("DropShadow", "false")
                                        ),
                                        new XElement("Name", ""),
                                        new XElement("Bmp", ""),
                                        new XElement("BmpMode", "0"),
                                        new XElement("Detached", "false"),
                                        new XElement("AllowMinimize", "false"),
                                        new XElement("AutoMinimize", "false"),
                                        new XElement("AllowMaximize", "false"),
                                        new XElement("AllowInfo", "false"),
                                        new XElement("PrintIcon", "false"),
                                        new XElement("CopyIcon", "false"),
                                        new XElement("XLIcon", "false"),
                                        new XElement("SearchIcon", "false"),
                                        new XElement("SelectPossibleIcon", "false"),
                                        new XElement("SelectExcludedIcon", "false"),
                                        new XElement("SelectAllIcon", "false"),
                                        new XElement("ClearIcon", "false"),
                                        new XElement("ClearOtherIcon", "false"),
                                        new XElement("LockIcon", "false"),
                                        new XElement("HelpText", ""),
                                        new XElement("ObjectId", ""),
                                        new XElement("ShrinkFrameToData", "true"),
                                        new XElement("ShowCaption", "false"),
                                        new XElement("MultiLine", "0"),
                                        new XElement("ActiveBgColor",
                                            new XElement("Mode", "COLORAREA_SOLID"),
                                            new XElement("Luminosity", "85"),
                                            new XElement("FillDirection", "COLORAREA_HORZ"),
                                            new XElement("FillPattern", "COLORAREA_PATTERN1"),
                                            new XElement("PrimaryCol",
                                                new XElement("Col",
                                                    new XElement("Red", "0"),
                                                    new XElement("Green", "106"),
                                                    new XElement("Blue", "157"),
                                                    new XElement("Alpha", "255")
                                                ),
                                                new XElement("IsCalculated", "false"),
                                                new XElement("ColorExpr", "")
                                            ),
                                        new XElement("SecondaryCol",
                                            new XElement("Col",
                                                new XElement("Red", "153"),
                                                new XElement("Green", "153"),
                                                new XElement("Blue", "153"),
                                                new XElement("Alpha", "255")
                                            ),
                                            new XElement("IsCalculated", "false"),
                                            new XElement("ColorExpr", "")
                                        ),
                                        new XElement("Alpha", "0")
                                    ),
                                    new XElement("BgColor",
                                        new XElement("Mode", "COLORAREA_SOLID"),
                                        new XElement("Luminosity", "85"),
                                        new XElement("FillDirection", "COLORAREA_HORZ"),
                                        new XElement("FillPattern", "COLORAREA_PATTERN1"),
                                        new XElement("PrimaryCol",
                                            new XElement("Col",
                                                new XElement("Red", "0"),
                                                new XElement("Green", "106"),
                                                new XElement("Blue", "157"),
                                                new XElement("Alpha", "255")
                                            ),
                                            new XElement("IsCalculated", "false"),
                                            new XElement("ColorExpr", "")
                                        ),
                                        new XElement("SecondaryCol",
                                            new XElement("Col",
                                                new XElement("Red", "200"),
                                                new XElement("Green", "200"),
                                                new XElement("Blue", "200"),
                                                new XElement("Alpha", "255")
                                            ),
                                            new XElement("IsCalculated", "false"),
                                            new XElement("ColorExpr", "")
                                        ),
                                        new XElement("Alpha", "0")
                                    ),
                                    new XElement("ActiveFgColor",
                                        new XElement("Mode", "COLORAREA_SOLID"),
                                        new XElement("Luminosity", "85"),
                                        new XElement("FillDirection", "COLORAREA_HORZ"),
                                        new XElement("FillPattern", "COLORAREA_PATTERN1"),
                                        new XElement("PrimaryCol",
                                            new XElement("Col",
                                                new XElement("Red", "255"),
                                                new XElement("Green", "255"),
                                                new XElement("Blue", "255"),
                                                new XElement("Alpha", "255")
                                            ),
                                            new XElement("IsCalculated", "false"),
                                            new XElement("ColorExpr", "")
                                        ),
                                        new XElement("SecondaryCol", ""),
                                        new XElement("Alpha", "255")
                                    ),
                                    new XElement("FgColor",
                                        new XElement("Mode", "COLORAREA_SOLID"),
                                        new XElement("Luminosity", "85"),
                                        new XElement("FillDirection", "COLORAREA_HORZ"),
                                        new XElement("FillPattern", "COLORAREA_PATTERN1"),
                                        new XElement("PrimaryCol",
                                            new XElement("Col",
                                                new XElement("Red", "255"),
                                                new XElement("Green", "255"),
                                                new XElement("Blue", "255"),
                                                new XElement("Alpha", "255")
                                            ),
                                            new XElement("IsCalculated", "false"),
                                            new XElement("ColorExpr", "")
                                        ),
                                        new XElement("SecondaryCol", ""),
                                        new XElement("Alpha", "255")
                                    ),
                                    new XElement("FixCorner", "true"),
                                    new XElement("FixCornerSize", "21"),
                                    new XElement("RelCornerSize", "4059000000000000"),
                                    new XElement("Power", "4000000000000000"),
                                    new XElement("RoundedShape", "true"),
                                    new XElement("EnableTopLeftRounded", "true"),
                                    new XElement("EnableTopRightRounded", "true"),
                                    new XElement("EnableBottomLeftRounded", "true"),
                                    new XElement("EnableBottomRightRounded", "true"),
                                    new XElement("Light", "3FEE666666666666"),
                                    new XElement("Dark", "3FC3333333333333"),
                                    new XElement("Rainbow", "false"),
                                    new XElement("TextAdjustHorizontal", "0"),
                                    new XElement("TextAdjustVertical", "4"),
                                    new XElement("Show",
                                        new XElement("Always", "false"),
                                        new XElement("Expression",
                                            new XElement("v", getCondition())
                                        )
                                    ),
                                    new XElement("AllowMoveSize", "true"),
                                    new XElement("AllowCopyClone", "true"),
                                    new XElement("dummy", ""),
                                    new XElement("CustomObjectGuid", ""),
                                    new XElement("CustomObjectProperties", ""),
                                    new XElement("ScrollBkgColor",
                                        new XElement("Col",
                                            new XElement("Red", "255"),
                                            new XElement("Green", "255"),
                                            new XElement("Blue", "255"),
                                            new XElement("Alpha", "255")
                                        ),
                                        new XElement("IsCalculated", "false"),
                                        new XElement("ColorExpr", "")
                                    ),
                                    new XElement("ScrollButtonColor",
                                        new XElement("Col",
                                            new XElement("Red", "192"),
                                            new XElement("Green", "192"),
                                            new XElement("Blue", "192"),
                                            new XElement("Alpha", "255")
                                        ),
                                        new XElement("IsCalculated", "false"),
                                        new XElement("ColorExpr", "")
                                    ),
                                    new XElement("ScrollStyle", "SCROLL_LIGHT"),
                                    new XElement("ScrollWidth", "38"),
                                    new XElement("AllowModify", "true"),
                                    new XElement("CopyImageIcon", "false"),
                                    new XElement("PreserveScrollPosition", "false"),
                                    new XElement("ShadowIntensity", "SHADOW_NONE"),
                                    new XElement("OnActivateActionItems", ""),
                                    new XElement("OnDeactivateActionItems", ""),
                                    new XElement("Extension", ""),
                                    new XElement("Background", ""),
                                    new XElement("MenuIcon", "false"),
                                    new XElement("ExtendedProperties", ""),
                                    new XElement("ExtendedValues", ""),
                                    new XElement("StateName", "")
                                ),
                                new XElement("LegacyHelpModeThatHasNoEffectWhatSoEver", "false"),
                                new XElement("TextAdjustHorizontal", "1"),
                                new XElement("TextAdjustVertical", "4"),
                                new XElement("TopMargin", "10"),
                                new XElement("BottomMargin", "10"),
                                new XElement("LeftMargin", "10"),
                                new XElement("RightMargin", "10"),
                                new XElement("TextColor",
                                    new XElement("Mode", "COLORAREA_SOLID"),
                                    new XElement("Luminosity", "85"),
                                    new XElement("FillDirection", "COLORAREA_HORZ"),
                                    new XElement("FillPattern", "COLORAREA_PATTERN1"),
                                    new XElement("PrimaryCol",
                                        new XElement("Col",
                                            new XElement("Red", "54"),
                                            new XElement("Green", "54"),
                                            new XElement("Blue", "54"),
                                            new XElement("Alpha", "255")
                                        ),
                                        new XElement("IsCalculated", "true"),
                                        new XElement("ColorExpr",
                                            new XElement("v", "vColor_Table_Text") //TODO. Get text color value as console param
                                        )
                                    ),
                                    new XElement("SecondaryCol", ""),
                                    new XElement("Alpha", "255")
                                ),
                                new XElement("BkgColor",
                                    new XElement("Mode", "COLORAREA_SOLID"),
                                    new XElement("Luminosity", "85"),
                                    new XElement("FillDirection", "COLORAREA_HORZ"),
                                    new XElement("FillPattern", "COLORAREA_PATTERN1"),
                                    new XElement("PrimaryCol",
                                        new XElement("Col",
                                            new XElement("Red", "214"),
                                            new XElement("Green", "231"),
                                            new XElement("Blue", "248"),
                                            new XElement("Alpha", "255")
                                        ),
                                        new XElement("IsCalculated", "true"),
                                        new XElement("ColorExpr",
                                            new XElement("v", bgColor)
                                        )
                                    ),
                                    new XElement("SecondaryCol", ""),
                                    new XElement("Alpha", "255")
                                ),
                                new XElement("BkgImageSettings", ""),
                                new XElement("BkgMode", "BKG_SOLID"),
                                new XElement("Bmp", ""),
                                new XElement("BkgAlpha", "255"),
                                new XElement("Text",
                                    new XElement("v", textVal) 
                                ),
                                new XElement("ImageRepresentation", ""),
                                new XElement("AllowVerticalScrollBar", "false"),
                                new XElement("AllowHorizontalScrollBar", "false"),
                                new XElement("ActionItems", "")
                            ),
                            new XElement("PrintSettings", ""),
                            new XElement("UserSized", "false")
                            );
            //TODO. Put this dogtail in wrapper. Without this bytes at file start Qlikview doesn't work.
            byte[] fileStart = new byte[] { 0xEF,0xBB,0xBF,0x0D,0x0A };
            using (BinaryWriter writer = new BinaryWriter(File.Open(Path.Combine(Program.path, objectID + ".xml"), FileMode.Create)))
            {
                writer.Write(fileStart);
            }
            File.AppendAllText(Path.Combine(Program.path, objectID + ".xml"), stubFile.ToString());
        }
    }
}
