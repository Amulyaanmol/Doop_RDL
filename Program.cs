using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;
namespace ConsoleApp
{
    public class TestModel
    {
        public string Name { get; set; }

        public IEnumerable<SharedData> GetSharedData { get; set; }
        public IEnumerable<Text> GetText { get; set; }
        public string TagName { get; set; }

        public string TagType { get; set; }
        public int NoLines { get; set; }

    }
    public class SharedData
    {
        public XName Type { get; set; }
        public string SharedDataSet { get; set; }
    }

    public class Text
    {
        public XName QueryType { get; set; }
        public XElement CmdType { get; set; }
        public string CmdText { get; set; }
    }

    public class TestModelList
    {
        public List<TestModel> testData { get; set; }
    }
    //install DocumentFormat.OpenXml nuget package
    public class Program
    {
        public static readonly string outputDirectory = @"C:\Users\91707\Desktop\MyOutput";
        public static readonly string workingDirectory = @"C:\Users\91707\Desktop\xml";
        public static int QueryStatus(TestModel textFileName, Text text)
        {
            var lv1 = textFileName;
            string outputFolder = outputDirectory;
            using TextWriter writer = new StreamWriter($"{outputFolder}\\{lv1.Name}.txt");
            writer.WriteLine(text.CmdText);
            writer.Close();
            var lineCount = File.ReadLines($"{outputFolder}\\{lv1.Name}.txt").Count();
            return lineCount;
        }
        public static IEnumerable<TestModel> FileParse(string files)
        {
            //IEnumerable<Level> lv1s = null
            XDocument xdoc = XDocument.Load(files);
            var ns = xdoc.Root.Attribute("xmlns").Value;
            //Adding the default namespace to process the RDL file
            XNamespace xns = XNamespace.Get(ns);
            var lv1s = from lv1 in xdoc.Descendants(xns + "DataSet")
                       select new TestModel
                       {
                           Name = lv1.Attribute("Name").Value,
                           GetSharedData = lv1.Descendants(xns + "SharedDataSet").
                                           Select(reference => new SharedData
                                           {
                                               Type = reference.Element(xns + "SharedDataSetReference").Parent.Name.LocalName,
                                               SharedDataSet = reference.Element(xns + "SharedDataSetReference").Value
                                           }),
                           GetText = lv1.Descendants(xns + "Query").
                                   Select(textValue => new Text
                                   {
                                       QueryType = textValue.Element(xns + "CommandText").Parent.Name.LocalName,
                                       CmdType = textValue.Element(xns + "CommandType"),
                                       CmdText = textValue.Element(xns + "CommandText").Value
                                   })
                       };
            return lv1s;
        }

        public static IEnumerable<TestModel> InnerSharedData(string fileName)
        {
            XDocument xdocs = XDocument.Load(fileName);
            var nsp = xdocs.Root.Attribute("xmlns").Value;
            XNamespace xnsp = XNamespace.Get(nsp);
            var lvl = from lv1 in xdocs.Descendants(xnsp + "DataSet")
                      select new TestModel
                      {
                          Name = Path.GetFileNameWithoutExtension(fileName),
                          GetSharedData = lv1.Descendants(xnsp + "SharedDataSet").
                                            Select(reference => new SharedData
                                            {
                                                Type = reference.Element(xnsp + "SharedDataSetReference").Parent.Name.LocalName,
                                                SharedDataSet = reference.Element(xnsp + "SharedDataSetReference").Value
                                            }),
                          GetText = lv1.Descendants(xnsp + "Query").
                  Select(textValue => new Text
                  {
                      QueryType = textValue.Element(xnsp + "CommandText").Parent.Name.LocalName,
                      CmdType = textValue.Element(xnsp + "CommandType"),
                      CmdText = textValue.Element(xnsp + "CommandText").Value
                  })
                      };
            return lvl;
        }

        public static void ListFilesInDirectory(string sourceFolder, string outputFolder)
        {
            using TextWriter logWriter = new StreamWriter($"{outputFolder}\\RDL_Log_File.txt");
            logWriter.WriteLine("Initializing the Directory.............\n");
            int queryCount = 0, queryTemp = 0, i=0, j = 0, sharedCount = 0, procedureCount = 0, procedureTemp = 0, totalTags = 0, functionTemp = 0, functionCount = 0;
            string extn = "", readLines = "";
            try
            {
                logWriter.WriteLine("Directory Started....\n");
                string[] filePaths = Directory.GetFiles(sourceFolder);
                if (filePaths.Length != 0)
                {
                    TestModelList tmList = new();
                    tmList.testData = new List<TestModel>();
                    foreach (string filePath in filePaths)
                    {
                        MatchCollection matchString;
                        FileInfo fi = new(filePath);
                        // Get file extension
                        extn = fi.Extension;
                        if (extn == ".rdl" || extn == ".xml" || extn == ".rsd")
                        {
                            var lv1s = FileParse(filePath);
                            foreach (var lv1 in lv1s)
                            {
                                TestModel tm = new();
                                tm.Name = Path.GetFileName(filePath);
                                foreach (var data in lv1.GetSharedData)
                                {

                                    const string rsdLocation = @"C:\Users\91707\Desktop\xml\SharedDataSetFolder";
                                    string r = data.SharedDataSet.Split('/').Last();
                                    string location = $"{rsdLocation}\\{r}.rsd";
                                    if (File.Exists(location))
                                    {
                                        tm.TagType = data.Type.LocalName;
                                        tm.TagName = data.SharedDataSet;
                                        var result = InnerSharedData(location);
                                        foreach (var l in result)
                                        {
                                            TestModel tm1 = new();
                                            tm1.Name = Path.GetFileName(l.Name + ".rsd");
                                            foreach (var dt in l.GetSharedData)
                                            {
                                                tm1.TagType = dt.Type.LocalName;
                                                tm1.TagName = dt.SharedDataSet;
                                                tmList.testData.Add(tm1);
                                            }
                                            foreach (var x in l.GetText)
                                            {
                                                var noLine = QueryStatus(l, x);
                                                tm1.TagName = l.Name;
                                                if (x.CmdType != null)
                                                {
                                                    tm1.TagType = x.CmdType.Value;
                                                    tm1.TagName = x.CmdText;
                                                }
                                                else
                                                {
                                                    tm1.TagType = x.QueryType.LocalName;
                                                    tm1.NoLines = noLine;
                                                }
                                                tmList.testData.Add(tm1);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        tm.TagType  = data.Type.LocalName;
                                        tm.TagName = data.SharedDataSet;
                                    }
                                    sharedCount += lv1.GetSharedData.Count();
                                    tmList.testData.Add(tm);
                                }
                                foreach (var x in lv1.GetText)
                                {
                                    var noLine = QueryStatus(lv1, x);
                                    tm.TagName = lv1.Name;
                                    if (x.CmdType != null)
                                    {
                                        tm.TagType = x.CmdType.Value;
                                        tm.TagName = x.CmdText;
                                        procedureCount += lv1.GetText.Count();
                                    }
                                    else
                                    {
                                        tm.TagType = x.QueryType.LocalName;
                                        queryCount += lv1.GetText.Count();
                                        tm.NoLines = noLine;
                                    }
                                    tmList.testData.Add(tm);
                                }
                            }
                        }
                        else
                        {
                            logWriter.WriteLine($"File - ({i}) [{Path.GetFileName(filePath)}] Causing the error \nInvalid xml file format found at File No : {i}\n\n");
                        }
                    StringBuilder sb = new();
                    XmlTextReader reader = new(filePath);
                    while (reader.Read())
                    {
                        if (reader.Name == "CommandText")
                        {
                            sb.AppendLine(reader.ReadString());
                        }
                    }
                    readLines = sb.ToString();
                    matchString = Regex.Matches(readLines, pattern: @"\b*dbo.*?(?=\()");
                    var readLs = string.Join(";", from Match match in matchString select match.Value.Trim());
                    IEnumerable<string> uniques = readLs.Split(';').Distinct();
                    foreach (var searchString in uniques)
                    {
                        TestModel tm2 = new();
                        tm2.Name = Path.GetFileName(filePath);
                        tm2.TagType = "Function";
                        tm2.TagName = searchString;
                        tmList.testData.Add(tm2);
                        Console.WriteLine(searchString);
                    }
                        logWriter.WriteLine($"({queryCount}), Query type tags and ({procedureCount}), Stored Procedure type tags has been found in current file\n");
                        logWriter.WriteLine(sharedCount + " Shared Data References found");
                        logWriter.WriteLine(functionCount + " Functions found");
                        functionTemp += functionCount;
                    queryTemp += queryCount;
                    procedureTemp += procedureCount;
                    totalTags += queryCount + procedureCount;
                    functionCount = 0;
                    queryCount = 0; procedureCount = 0;
                        logWriter.WriteLine(i + ", Files has been Parsed\n");

                    }
                    Program p = new();
                    p.CreateExcelFile(tmList, outputFolder);
                    logWriter.WriteLine($"({j}) .rdl files and ({i - j}) other extension type files has been Found");
                    logWriter.WriteLine($"--------------------------Parsing Completed\nTotal Query type tags : {queryTemp} \t Stored Procedure type tags {procedureTemp}\nTotal Files : {i} \t Total tags : {totalTags}");

                }
                else
                {
                    logWriter.WriteLine("Empty Files Please Check the given Directory Path");
                   
                }
            }
            catch (DirectoryNotFoundException directoryException)
            {
                logWriter.WriteLine(directoryException.ToString());
               
            }
            catch (ArgumentException nullArgumentException)
            {
                logWriter.WriteLine(nullArgumentException.ToString());
               
            }
        }

        static void Main(string[] args)
        {
            ListFilesInDirectory(workingDirectory, outputDirectory);
        }

        public void CreateExcelFile(TestModelList data, string OutPutFileDirectory)
        {
            var datetime = DateTime.Now.ToString().Replace("/", "_").Replace(":", "_");

            string fileFullname = Path.Combine(OutPutFileDirectory, "Output.xlsx");

            if (File.Exists(fileFullname))
            {
                fileFullname = Path.Combine(OutPutFileDirectory, "Output_" + datetime + ".xlsx");
            }

            using (SpreadsheetDocument package = SpreadsheetDocument.Create(fileFullname, SpreadsheetDocumentType.Workbook))
            {
                CreatePartsForExcel(package, data);
            }
        }

        private void CreatePartsForExcel(SpreadsheetDocument document, TestModelList data)
        {
            SheetData partSheetData = GenerateSheetdataForDetails(data);

            WorkbookPart workbookPart1 = document.AddWorkbookPart();
            GenerateWorkbookPartContent(workbookPart1);

            WorkbookStylesPart workbookStylesPart1 = workbookPart1.AddNewPart<WorkbookStylesPart>("rId3");
            GenerateWorkbookStylesPartContent(workbookStylesPart1);

            WorksheetPart worksheetPart1 = workbookPart1.AddNewPart<WorksheetPart>("rId1");
            GenerateWorksheetPartContent(worksheetPart1, partSheetData);
        }
        private void GenerateWorksheetPartContent(WorksheetPart worksheetPart1, SheetData sheetData1)
        {
            Worksheet worksheet = new() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            worksheet.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            SheetDimension sheetDimension1 = new SheetDimension() { Reference = "A1" };

            SheetViews sheetViews1 = new();

            SheetView sheetView1 = new() { TabSelected = true, WorkbookViewId = (UInt32Value)0U };
            Selection selection1 = new() { ActiveCell = "A1", SequenceOfReferences = new ListValue<StringValue>() { InnerText = "A1" } };

            sheetView1.Append(selection1);

            sheetViews1.Append(sheetView1);
            SheetFormatProperties sheetFormatProperties1 = new SheetFormatProperties() { DefaultRowHeight = 15D, DyDescent = 0.25D };

            PageMargins pageMargins1 = new PageMargins() { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
            worksheet.Append(sheetDimension1);
            worksheet.Append(sheetViews1);
            worksheet.Append(sheetFormatProperties1);
            worksheet.Append(sheetData1);
            worksheet.Append(pageMargins1);
            worksheetPart1.Worksheet = worksheet;
        }
        private void GenerateWorkbookStylesPartContent(WorkbookStylesPart workbookStylesPart1)
        {
            Stylesheet stylesheet1 = new() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            stylesheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            stylesheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            Fonts fonts1 = new() { Count = (UInt32Value)2U, KnownFonts = true };

            Font font1 = new();
            FontSize fontSize1 = new() { Val = 11D };
            Color color1 = new() { Theme = (UInt32Value)1U };
            FontName fontName1 = new() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering1 = new() { Val = 2 };
            FontScheme fontScheme1 = new() { Val = FontSchemeValues.Minor };

            font1.Append(fontSize1);
            font1.Append(color1);
            font1.Append(fontName1);
            font1.Append(fontFamilyNumbering1);
            font1.Append(fontScheme1);

            Font font2 = new Font();
            Bold bold1 = new Bold();
            FontSize fontSize2 = new FontSize() { Val = 11D };
            Color color2 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName2 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering2 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme2 = new FontScheme() { Val = FontSchemeValues.Minor };

            font2.Append(bold1);
            font2.Append(fontSize2);
            font2.Append(color2);
            font2.Append(fontName2);
            font2.Append(fontFamilyNumbering2);
            font2.Append(fontScheme2);

            fonts1.Append(font1);
            fonts1.Append(font2);

            Fills fills1 = new Fills() { Count = (UInt32Value)2U };

            Fill fill1 = new Fill();
            PatternFill patternFill1 = new PatternFill() { PatternType = PatternValues.None };

            fill1.Append(patternFill1);

            Fill fill2 = new Fill();
            PatternFill patternFill2 = new PatternFill() { PatternType = PatternValues.Gray125 };

            fill2.Append(patternFill2);

            fills1.Append(fill1);
            fills1.Append(fill2);

            Borders borders1 = new Borders() { Count = (UInt32Value)2U };

            Border border1 = new Border();
            LeftBorder leftBorder1 = new LeftBorder();
            RightBorder rightBorder1 = new RightBorder();
            TopBorder topBorder1 = new TopBorder();
            BottomBorder bottomBorder1 = new BottomBorder();
            DiagonalBorder diagonalBorder1 = new DiagonalBorder();

            border1.Append(leftBorder1);
            border1.Append(rightBorder1);
            border1.Append(topBorder1);
            border1.Append(bottomBorder1);
            border1.Append(diagonalBorder1);

            Border border2 = new Border();

            LeftBorder leftBorder2 = new LeftBorder() { Style = BorderStyleValues.Thin };
            Color color3 = new Color() { Indexed = (UInt32Value)64U };

            leftBorder2.Append(color3);

            RightBorder rightBorder2 = new RightBorder() { Style = BorderStyleValues.Thin };
            Color color4 = new Color() { Indexed = (UInt32Value)64U };

            rightBorder2.Append(color4);

            TopBorder topBorder2 = new TopBorder() { Style = BorderStyleValues.Thin };
            Color color5 = new Color() { Indexed = (UInt32Value)64U };

            topBorder2.Append(color5);

            BottomBorder bottomBorder2 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color6 = new Color() { Indexed = (UInt32Value)64U };

            bottomBorder2.Append(color6);
            DiagonalBorder diagonalBorder2 = new DiagonalBorder();

            border2.Append(leftBorder2);
            border2.Append(rightBorder2);
            border2.Append(topBorder2);
            border2.Append(bottomBorder2);
            border2.Append(diagonalBorder2);

            borders1.Append(border1);
            borders1.Append(border2);

            CellStyleFormats cellStyleFormats1 = new CellStyleFormats() { Count = (UInt32Value)1U };
            CellFormat cellFormat1 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };

            cellStyleFormats1.Append(cellFormat1);

            CellFormats cellFormats1 = new CellFormats() { Count = (UInt32Value)3U };
            CellFormat cellFormat2 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };
            CellFormat cellFormat3 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)0U, ApplyBorder = true };
            CellFormat cellFormat4 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true };

            cellFormats1.Append(cellFormat2);
            cellFormats1.Append(cellFormat3);
            cellFormats1.Append(cellFormat4);

            CellStyles cellStyles1 = new CellStyles() { Count = (UInt32Value)1U };
            CellStyle cellStyle1 = new CellStyle() { Name = "Normal", FormatId = (UInt32Value)0U, BuiltinId = (UInt32Value)0U };

            cellStyles1.Append(cellStyle1);
            DifferentialFormats differentialFormats1 = new DifferentialFormats() { Count = (UInt32Value)0U };
            TableStyles tableStyles1 = new TableStyles() { Count = (UInt32Value)0U, DefaultTableStyle = "TableStyleMedium2", DefaultPivotStyle = "PivotStyleLight16" };

            StylesheetExtensionList stylesheetExtensionList1 = new StylesheetExtensionList();

            StylesheetExtension stylesheetExtension1 = new StylesheetExtension() { Uri = "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" };
            stylesheetExtension1.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            X14.SlicerStyles slicerStyles1 = new X14.SlicerStyles() { DefaultSlicerStyle = "SlicerStyleLight1" };

            stylesheetExtension1.Append(slicerStyles1);

            StylesheetExtension stylesheetExtension2 = new StylesheetExtension() { Uri = "{9260A510-F301-46a8-8635-F512D64BE5F5}" };
            stylesheetExtension2.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
            X15.TimelineStyles timelineStyles1 = new X15.TimelineStyles() { DefaultTimelineStyle = "TimeSlicerStyleLight1" };

            stylesheetExtension2.Append(timelineStyles1);

            stylesheetExtensionList1.Append(stylesheetExtension1);
            stylesheetExtensionList1.Append(stylesheetExtension2);

            stylesheet1.Append(fonts1);
            stylesheet1.Append(fills1);
            stylesheet1.Append(borders1);
            stylesheet1.Append(cellStyleFormats1);
            stylesheet1.Append(cellFormats1);
            stylesheet1.Append(cellStyles1);
            stylesheet1.Append(differentialFormats1);
            stylesheet1.Append(tableStyles1);
            stylesheet1.Append(stylesheetExtensionList1);

            workbookStylesPart1.Stylesheet = stylesheet1;
        }
        private void GenerateWorkbookPartContent(WorkbookPart workbookPart1)
        {
            Workbook workbook1 = new Workbook();
            Sheets sheets1 = new Sheets();
            Sheet sheet1 = new Sheet() { Name = "Sheet1", SheetId = (UInt32Value)1U, Id = "rId1" };
            sheets1.Append(sheet1);
            workbook1.Append(sheets1);
            workbookPart1.Workbook = workbook1;
        }
        private SheetData GenerateSheetdataForDetails(TestModelList data)
        {
            SheetData sheetData1 = new SheetData();
            sheetData1.Append(CreateHeaderRowForExcel());

            foreach (TestModel taktTimemodel in data.testData)
            {
                Row partsRows = GenerateRowForChildPartDetail( taktTimemodel);
                sheetData1.Append(partsRows);
            }
            return sheetData1;
        }
        private Row CreateHeaderRowForExcel()
        {
            Row workRow = new Row();
            workRow.Append(CreateCell("File Name", 2U));
            workRow.Append(CreateCell("Tag Type", 2U));
            workRow.Append(CreateCell("Name", 2U));
            workRow.Append(CreateCell("No Of Lines", 2U));
            return workRow;
        }
        private Row GenerateRowForChildPartDetail(TestModel testmodel)
        {
            Row tRow = new Row();
            tRow.Append(CreateCell(testmodel.Name));
            tRow.Append(CreateCell(testmodel.TagType));
            tRow.Append(CreateCell(testmodel.TagName));
            tRow.Append(CreateCell(testmodel.NoLines.ToString()));
            return tRow;
        }

        private Cell CreateCell(string text)
        {
            Cell cell = new Cell();
            cell.StyleIndex = 1U;
            cell.DataType = ResolveCellDataTypeOnValue(text);
            cell.CellValue = new CellValue(text);
            return cell;
        }
        private Cell CreateCell(string text, uint styleIndex)
        {
            Cell cell = new Cell();
            cell.StyleIndex = styleIndex;
            cell.DataType = ResolveCellDataTypeOnValue(text);
            cell.CellValue = new CellValue(text);
            return cell;
        }
        private EnumValue<CellValues> ResolveCellDataTypeOnValue(string text)
        {
            int intVal;
            double doubleVal;
            if (int.TryParse(text, out intVal) || double.TryParse(text, out doubleVal))
            {
                return CellValues.Number;
            }
            else
            {
                return CellValues.String;
            }
        }
    }
}
