using System;
using Word = Microsoft.Office.Interop.Word;
using VBIDE = Microsoft.Vbe.Interop;

namespace ConsoleApp6
{
    class Program
    {
        static void Main(string[] args)
        {
            
            var wordApp = new Word.Application();
            wordApp.Documents.Add();//@"C:\Users\ASUS\Desktop\Титульник.docx");
            var doc = wordApp.Documents[1];
            VBIDE.VBComponent oModule = doc.VBProject.VBComponents.Add(VBIDE.vbext_ComponentType.vbext_ct_StdModule);

            try
            {
                string sCode =
                    //"public sub VBAMacro()\r\n" +
                    //"   msgbox \"VBA Macro called\"\r\n" +
                    //"end sub";
                "Sub Title()\r\n" +
                "Selection.PageSetup.TopMargin = CentimetersToPoints(1)\r\n" +
                "Selection.PageSetup.LeftMargin = CentimetersToPoints(2)\r\n" +
                "Selection.PageSetup.RightMargin = CentimetersToPoints(1)\r\n" +
                "Selection.Font.Name = \"Times New Roman\"\r\n" +
                "Selection.Font.Size = 12\r\n" +
                "Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter\r\n" +
                "Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle\r\n" +
                "Selection.ParagraphFormat.SpaceAfter = 0\r\n" +
                "Selection.Font.AllCaps = True\r\n" +
                "Selection.TypeText Text:= \"{ОБЪЕКТ}\"\r\n" +
                "Selection.TypeParagraph\r\n" +
                "Selection.TypeText Text:= \"{ЗАКАЗЧИК}\"\r\n" +
                "Selection.Font.AllCaps = False\r\n" +
                "Selection.TypeParagraph\r\n" +
                "Selection.TypeParagraph\r\n" +
                "Selection.TypeParagraph\r\n" +
                "Selection.TypeParagraph\r\n" +
                "Selection.TypeParagraph\r\n" +
                "Selection.TypeParagraph\r\n" +
                "Selection.TypeParagraph\r\n" +
                "Selection.TypeParagraph\r\n" +
                "Selection.TypeParagraph\r\n" +
                "Selection.TypeParagraph\r\n" +
                "Selection.TypeParagraph\r\n" +
                "Selection.TypeParagraph\r\n" +
                "Selection.TypeParagraph\r\n" +
                "Selection.TypeParagraph\r\n" +
                "Selection.Font.Size = 28\r\n" +
                "Selection.TypeText Text:= \"СДАТОЧНАЯ ДОКУМЕНТАЦИЯ\"\r\n" +
                "Selection.Font.Size = 12\r\n" +
                "Selection.TypeParagraph\r\n" +
                "Selection.TypeParagraph\r\n" +
                "Selection.TypeParagraph\r\n" +
                "Selection.TypeParagraph\r\n" +
                "Selection.TypeParagraph\r\n" +
                "Selection.TypeParagraph\r\n" +
                "Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft\r\n" +
                "ActiveDocument.Tables.Add Range:= Selection.Range, NumRows:= 2, NumColumns:= _\r\n" +
                    "2, DefaultTableBehavior:= wdWord9TableBehavior, AutoFitBehavior:= _\r\n" +
                    "wdAutoFitFixed\r\n" +
                "With Selection.Tables(1)\r\n" +
                    ".LeftPadding = CentimetersToPoints(0)\r\n" +
                    ".RightPadding = CentimetersToPoints(0)\r\n" +
                "End With\r\n" +
                "Selection.Tables(1).Columns(1).SetWidth ColumnWidth:= 42.8, RulerStyle:= _\r\n" +
                    "wdAdjustFirstColumn\r\n" +
                "Selection.Tables(1).Columns(2).SetWidth ColumnWidth:= 169.8, RulerStyle:= _\r\n" +
                    "wdAdjustFirstColumn\r\n" +
                "Selection.TypeText Text:= \"Проект:\"\r\n" +
                "Selection.Borders(wdBorderTop).LineStyle = wdLineStyleNone\r\n" +
                "Selection.Borders(wdBorderLeft).LineStyle = wdLineStyleNone\r\n" +
                "Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleNone\r\n" +
                "Selection.Borders(wdBorderRight).LineStyle = wdLineStyleNone\r\n" +
                "Selection.MoveRight Unit:= wdCharacter, Count:= 1\r\n" +
                "Selection.Font.Bold = wdToggle\r\n" +
                "Selection.Font.Italic = wdToggle\r\n" +
                "Selection.TypeText Text:= \"{Раздел проекта 1}\"\r\n" +
                "Selection.Borders(wdBorderTop).LineStyle = wdLineStyleNone\r\n" +
                "Selection.Borders(wdBorderRight).LineStyle = wdLineStyleNone\r\n" +
                "Selection.MoveDown Unit:= wdLine, Count:= 1\r\n" +
                "Selection.Borders(wdBorderLeft).LineStyle = wdLineStyleNone\r\n" +
                "Selection.Borders(wdBorderRight).LineStyle = wdLineStyleNone\r\n" +
                "Selection.MoveLeft Unit:= wdCharacter, Count:= 1\r\n" +
                "Selection.Borders(wdBorderLeft).LineStyle = wdLineStyleNone\r\n" +
                "Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleNone\r\n" +
                "Selection.MoveRight Unit:= wdCharacter, Count:= 1\r\n" +
                "Selection.Font.Bold = wdToggle\r\n" +
                "Selection.Font.Italic = wdToggle\r\n" +
                "Selection.TypeText Text:= \"{Раздел проекта 2}\"\r\n" +
                "Selection.MoveDown Unit:= wdLine, Count:= 1\r\n" +
                "Selection.TypeParagraph\r\n" +
                "Selection.TypeParagraph\r\n" +
                "Selection.TypeParagraph\r\n" +
                "Selection.TypeParagraph\r\n" +
                "Selection.TypeParagraph\r\n" +
                "ActiveDocument.Tables.Add Range:= Selection.Range, NumRows:= 5, NumColumns:= _\r\n" +
                    "2, DefaultTableBehavior:= wdWord9TableBehavior, AutoFitBehavior:= _\r\n" +
                    "wdAutoFitFixed\r\n" +
                "With Selection.Tables(1)\r\n" +
                    ".LeftPadding = CentimetersToPoints(0)\r\n" +
                    ".RightPadding = CentimetersToPoints(0)\r\n" +
                "End With\r\n" +
                "Selection.Tables(1).Columns(1).SetWidth ColumnWidth:= 56.8, RulerStyle:= _\r\n" +
                    "wdAdjustFirstColumn\r\n" +
                "Selection.Tables(1).Columns(2).SetWidth ColumnWidth:= 453.7, RulerStyle:= _\r\n" +
                    "wdAdjustFirstColumn\r\n" +
                "Selection.Tables(1).Cell(1, 2).Merge Selection.Tables(1).Cell(4, 2)\r\n" +
                "Selection.TypeText Text:= \"Объект:\"\r\n" +
                "Selection.Borders(wdBorderTop).LineStyle = wdLineStyleNone\r\n" +
                "Selection.Borders(wdBorderLeft).LineStyle = wdLineStyleNone\r\n" +
                "Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleNone\r\n" +
                "Selection.Borders(wdBorderRight).LineStyle = wdLineStyleNone\r\n" +
                "Selection.MoveRight Unit:= wdCharacter, Count:= 1\r\n" +
                "Selection.Font.Bold = wdToggle\r\n" +
                "Selection.Font.Italic = wdToggle\r\n" +
                "Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify\r\n" +
                "Selection.TypeText Text:= \" {Полное наименование проекта}\"\r\n" +
                "Selection.Borders(wdBorderTop).LineStyle = wdLineStyleNone\r\n" +
                "Selection.Borders(wdBorderLeft).LineStyle = wdLineStyleNone\r\n" +
                "Selection.Borders(wdBorderRight).LineStyle = wdLineStyleNone\r\n" +
                "With ActiveDocument.Shapes.AddLine(114, 448, 567, 448).Line\r\n" +
                    ".ForeColor.RGB = RGB(0, 0, 0)\r\n" +
                "End With\r\n" +
                "With ActiveDocument.Shapes.AddLine(114, 462, 567, 462).Line\r\n" +
                    ".ForeColor.RGB = RGB(0, 0, 0)\r\n" +
                "End With\r\n" +
                "With ActiveDocument.Shapes.AddLine(114, 476.5, 567, 476.5).Line\r\n" +
                    ".ForeColor.RGB = RGB(0, 0, 0)\r\n" +
                "End With\r\n" +
                "Selection.MoveDown Unit:= wdLine, Count:= 1\r\n" +
                "Selection.Font.Bold = wdToggle\r\n" +
                "Selection.Font.Italic = wdToggle\r\n" +
                "Selection.TypeText Text:= \"{адрес объекта}\"\r\n" +
                "Selection.Borders(wdBorderLeft).LineStyle = wdLineStyleNone\r\n" +
                "Selection.Borders(wdBorderRight).LineStyle = wdLineStyleNone\r\n" +
                "Selection.MoveLeft Unit:= wdCharacter, Count:= 16\r\n" +
                "Selection.Borders(wdBorderLeft).LineStyle = wdLineStyleNone\r\n" +
                "Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleNone\r\n" +
                "Selection.TypeText Text:= \"по адресу:\"\r\n" +
                "Selection.MoveUp Unit:= wdLine, Count:= 1\r\n" +
                "Selection.Borders(wdBorderLeft).LineStyle = wdLineStyleNone\r\n" +
                "Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleNone\r\n" +
                "Selection.MoveUp Unit:= wdLine, Count:= 1\r\n" +
                "Selection.Borders(wdBorderLeft).LineStyle = wdLineStyleNone\r\n" +
                "Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleNone\r\n" +
                "Selection.MoveUp Unit:= wdLine, Count:= 1\r\n" +
                "Selection.Borders(wdBorderLeft).LineStyle = wdLineStyleNone\r\n" +
                "Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleNone\r\n" +
                "Selection.MoveDown Unit:= wdLine, Count:= 4\r\n" +
                "Selection.Font.Size = 8\r\n" +
                "Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter\r\n" +
                "Selection.TypeText Text:= \"(адрес объекта)\"\r\n" +
                "Selection.TypeParagraph\r\n" +
                "Selection.Font.Size = 12\r\n" +
                "Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft\r\n" +
                "Selection.TypeParagraph\r\n" +
                "Selection.TypeParagraph\r\n" +
                "Selection.TypeParagraph\r\n" +
                "Selection.TypeParagraph\r\n" +
                "Selection.TypeParagraph\r\n" +
                "Selection.TypeParagraph\r\n" +
                "Selection.TypeParagraph\r\n" +
                "ActiveDocument.Tables.Add Range:= Selection.Range, NumRows:= 1, NumColumns:= _\r\n" +
                    "2, DefaultTableBehavior:= wdWord9TableBehavior, AutoFitBehavior:= _\r\n" +
                    "wdAutoFitFixed\r\n" +
                "With Selection.Tables(1)\r\n" +
                    ".LeftPadding = CentimetersToPoints(0)\r\n" +
                    ".RightPadding = CentimetersToPoints(0)\r\n" +
                "End With\r\n" +
                "Selection.Tables(1).Columns(1).SetWidth ColumnWidth:= 71#, RulerStyle:= _\r\n" +
                    "wdAdjustFirstColumn\r\n" +
                "Selection.Tables(1).Columns(2).SetWidth ColumnWidth:= 439.4, RulerStyle:= _\r\n" +
                    "wdAdjustFirstColumn\r\n" +
                "Selection.TypeText Text:= \"Исполнитель:\"\r\n" +
                "Selection.Borders(wdBorderTop).LineStyle = wdLineStyleNone\r\n" +
                "Selection.Borders(wdBorderLeft).LineStyle = wdLineStyleNone\r\n" +
                "Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleNone\r\n" +
                "Selection.Borders(wdBorderRight).LineStyle = wdLineStyleNone\r\n" +
                "Selection.MoveRight Unit:= wdCharacter, Count:= 1\r\n" +
                "Selection.Font.Bold = wdToggle\r\n" +
                "Selection.Font.Italic = wdToggle\r\n" +
                "Selection.TypeText Text:= \" {должность} ЗАО \"\"ГК \"\"ТЭКС-Автоматик\"\" {Ф.И.О.}\"\r\n" +
                "Selection.Borders(wdBorderTop).LineStyle = wdLineStyleNone\r\n" +
                "Selection.Borders(wdBorderLeft).LineStyle = wdLineStyleNone\r\n" +
                "Selection.Borders(wdBorderRight).LineStyle = wdLineStyleNone\r\n" +
                "Selection.MoveDown Unit:= wdLine, Count:= 1\r\n" +
                "Selection.Font.Size = 8\r\n" +
                "Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter\r\n" +
                "Selection.TypeText Text:= \"(должность, наименование монтажной организации, Ф.И.О.)\"\r\n" +
                "Selection.TypeParagraph\r\n" +
                "Selection.Font.Size = 12\r\n" +
                "Selection.TypeParagraph\r\n" +
                "Selection.TypeParagraph\r\n" +
                "Selection.TypeParagraph\r\n" +
                "Selection.TypeParagraph\r\n" +
                "Selection.TypeParagraph\r\n" +
                "Selection.TypeParagraph\r\n" +
                "Selection.TypeParagraph\r\n" +
                "Selection.TypeParagraph\r\n" +
                "Selection.TypeParagraph\r\n" +
                "Selection.TypeText Text:= \"20{год} г.\"\r\n" +
                "End Sub\r\n";
                // Add the VBA macro to the new code module.
                oModule.CodeModule.AddFromString(sCode);
                RunMacro(wordApp, new Object[] { "Title" });
                wordApp.Visible = false;
                doc.SaveAs2(@"C:\Users\ASUS\Desktop\Титульник.docx");
                doc.Close();
                wordApp.Quit();
            }

            catch
            {
                Console.WriteLine("Произошла ошибка");
            }
        }
        public static void RunMacro(object oApp, object[] oRunArgs)
        {
            oApp.GetType().InvokeMember("Run",
                System.Reflection.BindingFlags.Default |
                System.Reflection.BindingFlags.InvokeMethod,
                null, oApp, oRunArgs);
        }
    } 
}
