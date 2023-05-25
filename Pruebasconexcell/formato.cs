using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Color = DocumentFormat.OpenXml.Spreadsheet.Color;

namespace tablitas
{
    public class formato
    {
        
        public static void formatear(string filename, int columnas, string color, string condicion, uint id)
            
        {
            Dictionary<int, string> letras = new Dictionary<int, string>()
        {
        {0, ""},
        {1, "A"}, {2, "B"}, {3, "C"}, {4, "D"}, {5, "E"}, {6, "F"}, {7, "G"}, {8, "H"}, {9, "I"}, {10, "J"},
        {11, "K"}, {12, "L"}, {13, "M"}, {14, "N"}, {15, "O"}, {16, "P"}, {17, "Q"}, {18, "R"}, {19, "S"}, {20, "T"},
        {21, "U"}, {22, "V"}, {23, "W"}, {24, "X"}, {25, "Y"}, {26, "Z"}
        };

            int inicio = (columnas + 5) / 2;
            int fin = columnas - 1;

            string iniciostr = letras[inicio];
            string finstr = letras[fin];

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(filename, true))
            {
                WorkbookPart workbookPart = document.WorkbookPart;
                //get the correct sheet
                Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().First();
                WorksheetPart worksheetPart = workbookPart.GetPartById(sheet.Id) as WorksheetPart;
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                //grab the stylesPart so we can add the style to apply (create one if one doesn't already exist)
                WorkbookStylesPart stylesPart = document.WorkbookPart.GetPartsOfType<WorkbookStylesPart>().FirstOrDefault();
                if (stylesPart == null)
                {
                    stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                    stylesPart.Stylesheet = new Stylesheet();
                }

                //create a fills object to hold the background colour we're going to apply
                Fills fills = new Fills() { Count = 1U };

                //grab the differential formats part so we can add the style to apply (create one if one doesn't already exist)
                bool addDifferentialFormats = false;
                DifferentialFormats differentialFormats = stylesPart.Stylesheet.GetFirstChild<DifferentialFormats>();
                if (differentialFormats == null)
                {
                    differentialFormats = new DifferentialFormats() { Count = 1U };
                    addDifferentialFormats = true;
                }

                //create the conditional format reference
                ConditionalFormatting conditionalFormatting = new ConditionalFormatting()
                {
                    SequenceOfReferences = new ListValue<StringValue>()
                    {
                        InnerText = $"${iniciostr}$2:${finstr}$1048576"
                    }
                };

                //create a style to assign to the conditional format
                DifferentialFormat differentialFormat = new DifferentialFormat();
                Fill fill = new Fill();
                PatternFill patternFill = new PatternFill();
                BackgroundColor backgroundColor = new BackgroundColor() { Rgb = new HexBinaryValue() { Value = color } };
                patternFill.Append(backgroundColor);
                fill.Append(patternFill);
                differentialFormat.Append(fill);
                var a = differentialFormats.Count;
                differentialFormats.Append(differentialFormat);

                //create the formula
                Formula formula1 = new Formula();
                formula1.Text = $"={iniciostr}2{condicion}" /*"INDIRECT(\"D\"&ROW())=\"Disapproved\""*/;

                //create a new conditional formatting rule with a type of Expression
                ConditionalFormattingRule conditionalFormattingRule = new ConditionalFormattingRule()
                {
                    Type = ConditionalFormatValues.Expression,
                    FormatId = id,
                    Priority = 1
                };

                //append the formula to the rule
                conditionalFormattingRule.Append(formula1);

                //append th formatting rule to the formatting collection
                conditionalFormatting.Append(conditionalFormattingRule);

                //add the formatting collection to the worksheet
                //note the ordering is important; there are other elements that should be checked for here really.
                //See the spec for all of them and see https://stackoverflow.com/questions/25398450/why-appending-autofilter-corrupts-my-excel-file-in-this-example/25410242#25410242
                //for more details on ordering
                PageMargins margins = worksheetPart.Worksheet.GetFirstChild<PageMargins>();
                if (margins != null)
                    worksheetPart.Worksheet.InsertBefore(conditionalFormatting, margins);
                else
                    worksheetPart.Worksheet.Append(conditionalFormatting);

                //add the differential formats to the stylesheet if it didn't already exist
                if (addDifferentialFormats)
                    stylesPart.Stylesheet.Append(differentialFormats);

                workbookPart.Workbook.Save();
            }

            
        }
    }
}
