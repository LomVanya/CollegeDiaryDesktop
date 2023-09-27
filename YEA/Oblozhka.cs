using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;
using Shape = Microsoft.Office.Interop.Word.Shape;
using Microsoft.Office.Core;


namespace YEA
{
    public  class Oblozhka
    {
        private string praktika;
        private string name;
        private string otchestvo;
        private string surname;
        private string group;
        private string specialization;
        private string cvalification;
        private string prepod;
        private string prepodTwo;




        public Oblozhka(string praktika, string name, string otchestvo, string surname, string group, string specialization, string cvalification, string prepod, string prepodTwo)
        {
            this.praktika = praktika;
            this.name = name;
            this.otchestvo = otchestvo;
            this.surname = surname;
            this.group = group;
            this.specialization = specialization;
            this.cvalification = cvalification;
            this.prepod = prepod;
            this.prepodTwo = prepodTwo; 

        }

        public Oblozhka()
        {

        }

        public Oblozhka(string surname)
        {
            this.surname = surname;
        }

        public string Praktika
        {
            get { return praktika; }
            set { praktika = value; }
        }

        public string Name
        {
            get { return name; }
            set { name = value; }
        }

        public string Otchestvo
        {
            get { return otchestvo; }
            set { otchestvo = value; }
        }
        public string Surname
        {
            get { return surname; }
            set { surname = value; }
        }

        public string Group
        {
            get { return group; }
            set { group = value; }
        }
        public string Specialization
        {
            get { return specialization; }
            set { specialization = value; }
        }

        public string Cvalification
        {
            get { return cvalification; }
            set { cvalification = value; }
        }
        public string Prepod
        {
            get { return prepod; }
            set { prepod = value; }
        }

        public string PrepodTwo
        {
            get { return prepodTwo; }
            set { prepodTwo = value; }
        }



        public void CreateDocument()
        {

            Word.Application wordApp = new Word.Application();
            Word.Document wordDoc = wordApp.Documents.Add();

            wordDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;
            Word.Section section = wordDoc.Sections[1];
            section.PageSetup.TopMargin = wordApp.CentimetersToPoints(1.6f);
            section.PageSetup.LeftMargin = wordApp.CentimetersToPoints(1.0f);

            Word.Paragraph institutionParagraph = wordDoc.Content.Paragraphs.Add();
            institutionParagraph.Range.Text = "Частное учреждение образования";
            institutionParagraph.Range.Font.Size = 10;
            institutionParagraph.Range.Font.Name = "Times New Roman";
            institutionParagraph.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            institutionParagraph.Range.ParagraphFormat.LeftIndent = wordApp.CentimetersToPoints(16.5f);
            institutionParagraph.Range.ParagraphFormat.SpaceAfter = 0;
            institutionParagraph.Range.InsertParagraphAfter();

            Word.Paragraph collegeParagraph = wordDoc.Content.Paragraphs.Add();
            collegeParagraph.Range.Text = "«КОЛЛЕДЖ БИЗНЕСА И ПРАВА»";
            collegeParagraph.Range.Font.Size = 10;
            collegeParagraph.Range.Font.Name = "Times New Roman";
            collegeParagraph.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            collegeParagraph.Range.ParagraphFormat.LeftIndent = wordApp.CentimetersToPoints(16.5f);
            collegeParagraph.Range.InsertParagraphAfter();

            for (int i = 0; i < 5; i++)
            {
                Word.Paragraph emptyParagraph = wordDoc.Content.Paragraphs.Add();
                emptyParagraph.Range.InsertParagraphAfter();
            }

            Word.Paragraph headerParagraph = wordDoc.Content.Paragraphs.Add();
            headerParagraph.Range.Text = "ДНЕВНИК";
            headerParagraph.Range.Font.Size = 14;
            headerParagraph.Range.Font.Name = "Times New Roman";
            headerParagraph.Range.Font.Bold = 1;
            headerParagraph.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            headerParagraph.Range.ParagraphFormat.SpaceAfter = 8;
            headerParagraph.Range.ParagraphFormat.LeftIndent = wordApp.CentimetersToPoints(16.5f);
            headerParagraph.Range.InsertParagraphAfter();

            Word.Paragraph praktikaparaghraph = wordDoc.Content.Paragraphs.Add();
            praktikaparaghraph.Range.Text = "ПРОХОЖДЕНИЯ ПРАКТИКИ";
            praktikaparaghraph.Range.Font.Size = 10;
            praktikaparaghraph.Range.Font.Name = "Times New Roman";
            praktikaparaghraph.Range.Font.Bold = 1;
            praktikaparaghraph.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            praktikaparaghraph.Range.ParagraphFormat.LeftIndent = wordApp.CentimetersToPoints(16.5f);
            praktikaparaghraph.Range.ParagraphFormat.SpaceAfter = 38;
            praktikaparaghraph.Range.InsertParagraphAfter();


            float lineWidth = 4.8f;
            Shape lineShape = wordDoc.Shapes.AddLine(
                wordApp.InchesToPoints(6.6f),
                wordApp.InchesToPoints(3.74f),
                wordApp.InchesToPoints(lineWidth + 6.6f),
                wordApp.InchesToPoints(3.74f)
            );
           
            lineShape.Line.Weight = 0f;


            Word.Paragraph nameofp = wordDoc.Content.Paragraphs.Add();
            nameofp.Range.Text = "(наименование практики)";
            nameofp.Range.Font.Size = 8;
            nameofp.Range.Font.Name = "Times New Roman";
            nameofp.Range.Font.Bold = 0;
            nameofp.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            nameofp.Range.ParagraphFormat.LeftIndent = wordApp.CentimetersToPoints(16.5f);
            nameofp.Range.ParagraphFormat.SpaceAfter = 18;
            nameofp.Range.InsertParagraphAfter();


            Word.Paragraph obuch = wordDoc.Content.Paragraphs.Add();
            obuch.Range.Text = "обучающегося";
            obuch.Range.Font.Size = 10;
            obuch.Range.Font.Name = "Times New Roman";
            obuch.Range.Font.Bold = 0;
            obuch.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            obuch.Range.ParagraphFormat.LeftIndent = wordApp.CentimetersToPoints(6.5f);
            obuch.Range.ParagraphFormat.SpaceAfter = 0;
            obuch.Range.InsertParagraphAfter();


            float lineWidth1 = 3.9f;
            Shape lineShape1 = wordDoc.Shapes.AddLine(
                wordApp.InchesToPoints(7.5f),
                wordApp.InchesToPoints(4.27f),
                wordApp.InchesToPoints(lineWidth1 + 7.5f),
                wordApp.InchesToPoints(4.27f)
            );
   
            lineShape1.Line.Weight = 0f;


            Word.Paragraph nameYOu = wordDoc.Content.Paragraphs.Add();
            nameYOu.Range.Text = "(фамилия, собственное имя, отчество (если таковое имеется)";
            nameYOu.Range.Font.Size = 8;
            nameYOu.Range.Font.Name = "Times New Roman";
            nameYOu.Range.Font.Bold = 0;
            nameYOu.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            nameYOu.Range.ParagraphFormat.LeftIndent = wordApp.CentimetersToPoints(18.5f);
            nameYOu.Range.ParagraphFormat.SpaceAfter = 16;
            nameYOu.Range.InsertParagraphAfter();


            Word.Paragraph nameSpecial = wordDoc.Content.Paragraphs.Add();
            nameSpecial.Range.Text = "Специальность";
            nameSpecial.Range.Font.Size = 10;
            nameSpecial.Range.Font.Name = "Times New Roman";
            nameSpecial.Range.Font.Bold = 0;
            nameSpecial.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            nameSpecial.Range.ParagraphFormat.LeftIndent = wordApp.CentimetersToPoints(6.55f);
            nameSpecial.Range.ParagraphFormat.SpaceAfter = 28;
            nameSpecial.Range.InsertParagraphAfter();

            float lineWidth2 = 3.9f;
            Shape lineShape2 = wordDoc.Shapes.AddLine(
                wordApp.InchesToPoints(7.5f),
                wordApp.InchesToPoints(4.80f),
                wordApp.InchesToPoints(lineWidth2 + 7.5f),
                wordApp.InchesToPoints(4.80f)
            );
     
            lineShape1.Line.Weight = 0f;

            float lineWidth3 = 4.8f;
            Shape lineShape3 = wordDoc.Shapes.AddLine(
                wordApp.InchesToPoints(6.6f),
                wordApp.InchesToPoints(5.09f),
                wordApp.InchesToPoints(lineWidth3 + 6.6f),
                wordApp.InchesToPoints(5.09f)
            );
   
            lineShape.Line.Weight = 0f;


            Word.Paragraph nameCvali = wordDoc.Content.Paragraphs.Add();
            nameCvali.Range.Text = "Квалификация";
            nameCvali.Range.Font.Size = 10;
            nameCvali.Range.Font.Name = "Times New Roman";
            nameCvali.Range.Font.Bold = 0;
            nameCvali.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            nameCvali.Range.ParagraphFormat.LeftIndent = wordApp.CentimetersToPoints(6.5f);
            nameCvali.Range.ParagraphFormat.SpaceAfter = 6;
            nameCvali.Range.InsertParagraphAfter();


            float lineWidth4 = 3.9f;
            Shape lineShape4 = wordDoc.Shapes.AddLine(
                wordApp.InchesToPoints(7.5f),
                wordApp.InchesToPoints(5.36f),
                wordApp.InchesToPoints(lineWidth4 + 7.5f),
                wordApp.InchesToPoints(5.36f)
            );
 
            lineShape1.Line.Weight = 0f;


            Shape labelShape = wordDoc.Shapes.AddLabel(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 50, 50, 100, 50);


            labelShape.TextFrame.TextRange.Text = praktika;

            labelShape.Height = 20;
            labelShape.Width = 350;

            labelShape.TextFrame.TextRange.Font.Name = "Times New Roman";
            labelShape.TextFrame.TextRange.Font.Size = 10;


            labelShape.Left = 450;
            labelShape.Top = 210;


            labelShape.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;

            Word.Paragraph nameGr = wordDoc.Content.Paragraphs.Add();
            nameGr.Range.Text = "Группа№";
            nameGr.Range.Font.Size = 10;
            nameGr.Range.Font.Name = "Times New Roman";
            nameGr.Range.Font.Bold = 0;
            nameGr.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            nameGr.Range.ParagraphFormat.LeftIndent = wordApp.CentimetersToPoints(5.75f);
            nameGr.Range.ParagraphFormat.SpaceAfter = 6;
            nameGr.Range.InsertParagraphAfter();

            float lineWidth5 = 3.9f;
            Shape lineShape5 = wordDoc.Shapes.AddLine(
                wordApp.InchesToPoints(7.2f),
                wordApp.InchesToPoints(5.61f),
                wordApp.InchesToPoints(lineWidth5 + 7.5f),
                wordApp.InchesToPoints(5.61f)
            );

            lineShape1.Line.Weight = 0f;

            Word.Paragraph nameNezn1 = wordDoc.Content.Paragraphs.Add();
            nameNezn1.Range.Text = "Руководитель практики от учреждения образования:";
            nameNezn1.Range.Font.Size = 10;
            nameNezn1.Range.Font.Name = "Times New Roman";
            nameNezn1.Range.Font.Bold = 0;
            nameNezn1.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            nameNezn1.Range.ParagraphFormat.LeftIndent = wordApp.CentimetersToPoints(12.25f);
            nameNezn1.Range.ParagraphFormat.SpaceAfter = 16;
            nameNezn1.Range.InsertParagraphAfter();


            float lineWidth6 = 4.8f;
            Shape lineShape6 = wordDoc.Shapes.AddLine(
                wordApp.InchesToPoints(6.6f),
                wordApp.InchesToPoints(6.12f),
                wordApp.InchesToPoints(lineWidth6 + 6.6f),
                wordApp.InchesToPoints(6.12f)
            );

            lineShape.Line.Weight = 0f;

            Word.Paragraph nameNezn2 = wordDoc.Content.Paragraphs.Add();
            nameNezn2.Range.Text = "(инициалы, фамилия)";
            nameNezn2.Range.Font.Size = 8;
            nameNezn2.Range.Font.Name = "Times New Roman";
            nameNezn2.Range.Font.Bold = 0;
            nameNezn2.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            nameNezn2.Range.ParagraphFormat.LeftIndent = wordApp.CentimetersToPoints(16.3f);
            nameNezn2.Range.ParagraphFormat.SpaceAfter = 6;
            nameNezn2.Range.InsertParagraphAfter();

            Word.Paragraph nameNezn3 = wordDoc.Content.Paragraphs.Add();
            nameNezn3.Range.Text = "Руководитель практики от организации <*>:";
            nameNezn3.Range.Font.Size = 10;
            nameNezn3.Range.Font.Name = "Times New Roman";
            nameNezn3.Range.Font.Bold = 0;
            nameNezn3.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            nameNezn3.Range.ParagraphFormat.LeftIndent = wordApp.CentimetersToPoints(11f);
            nameNezn3.Range.ParagraphFormat.SpaceAfter = 16;
            nameNezn3.Range.InsertParagraphAfter();



            float lineWidth7 = 4.8f;
            Shape lineShape7 = wordDoc.Shapes.AddLine(
                wordApp.InchesToPoints(6.6f),
                wordApp.InchesToPoints(6.74f),
                wordApp.InchesToPoints(lineWidth7 + 6.6f),
                wordApp.InchesToPoints(6.74f)
            );

            lineShape.Line.Weight = 0f;

            Word.Paragraph nameNezn4 = wordDoc.Content.Paragraphs.Add();
            nameNezn4.Range.Text = "(инициалы, фамилия)";
            nameNezn4.Range.Font.Size = 8;
            nameNezn4.Range.Font.Name = "Times New Roman";
            nameNezn4.Range.Font.Bold = 0;
            nameNezn4.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            nameNezn4.Range.ParagraphFormat.LeftIndent = wordApp.CentimetersToPoints(16.3f);
            nameNezn4.Range.ParagraphFormat.SpaceAfter = 6;
            nameNezn4.Range.InsertParagraphAfter();




            Shape labelShape2 = wordDoc.Shapes.AddLabel(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 50, 50, 100, 50);


            labelShape2.TextFrame.TextRange.Text = surname + " " + name + " " + otchestvo;

            labelShape2.Height = 20;
            labelShape2.Width = 350;

            labelShape2.TextFrame.TextRange.Font.Name = "Times New Roman";
            labelShape2.TextFrame.TextRange.Font.Size = 10;


            labelShape2.Left = 510;
            labelShape2.Top = 247;

            labelShape2.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;

            Shape labelShape3 = wordDoc.Shapes.AddLabel(MsoTextOrientation.msoTextOrientationHorizontal, 50, 50, 100, 50);


        

            labelShape3.TextFrame.TextRange.Text = specialization;

            labelShape3.Height = 20;
            labelShape3.Width = 350;

            labelShape3.TextFrame.TextRange.Font.Name = "Times New Roman";
            labelShape3.TextFrame.TextRange.Font.Size = 10;


            labelShape3.Left = 510;
            labelShape3.Top = 286;


            labelShape3.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;


            Shape labelShape4 = wordDoc.Shapes.AddLabel(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 50, 50, 100, 50);


            labelShape4.TextFrame.TextRange.Text = cvalification;

            labelShape4.Height = 20;
            labelShape4.Width = 350;

            labelShape4.TextFrame.TextRange.Font.Name = "Times New Roman";
            labelShape4.TextFrame.TextRange.Font.Size = 10;


            labelShape4.Left = 510;
            labelShape4.Top = 326;


            labelShape4.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;


            Shape labelShape5 = wordDoc.Shapes.AddLabel(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 50, 50, 100, 50);


            labelShape5.TextFrame.TextRange.Text = group;

            labelShape5.Height = 20;
            labelShape5.Width = 350;

            labelShape5.TextFrame.TextRange.Font.Name = "Times New Roman";
            labelShape5.TextFrame.TextRange.Font.Size = 10;


            labelShape5.Left = 487;
            labelShape5.Top = 344;

            labelShape5.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;



            Shape labelShape6 = wordDoc.Shapes.AddLabel(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 50, 50, 100, 50);


            labelShape6.TextFrame.TextRange.Text = prepod;

            labelShape6.Height = 20;
            labelShape6.Width = 350;

            labelShape6.TextFrame.TextRange.Font.Name = "Times New Roman";
            labelShape6.TextFrame.TextRange.Font.Size = 10;


            labelShape6.Left = 450;
            labelShape6.Top = 381;

            labelShape6.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;


            Shape labelShape7 = wordDoc.Shapes.AddLabel(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 50, 50, 100, 50);

            labelShape7.TextFrame.TextRange.Text = prepodTwo;

            labelShape7.Height = 20;
            labelShape7.Width = 341;

            labelShape7.TextFrame.TextRange.Font.Name = "Times New Roman";
            labelShape7.TextFrame.TextRange.Font.Size = 10;


            labelShape7.Left = 450;
            labelShape7.Top = 425;

            labelShape7.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;




            lineShape.Line.ForeColor.RGB = 0x000000;
            lineShape1.Line.ForeColor.RGB = 0x000000;
            lineShape2.Line.ForeColor.RGB = 0x000000;
            lineShape3.Line.ForeColor.RGB = 0x000000;
            lineShape4.Line.ForeColor.RGB = 0x000000;
            lineShape5.Line.ForeColor.RGB = 0x000000;
            lineShape6.Line.ForeColor.RGB = 0x000000;
            lineShape7.Line.ForeColor.RGB = 0x000000;



            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "Документ Word (*.docx)|*.docx";
            saveFileDialog1.Title = "Сохранить документ Word";
            saveFileDialog1.ShowDialog();

            if (saveFileDialog1.FileName != "")
            {
                wordDoc.SaveAs2(saveFileDialog1.FileName);
            }


            wordDoc.Close();
            wordApp.Quit();




        }

        public void SignsForming()
        {
            Word.Application wordApp = new Word.Application();
            Word.Document wordDoc = wordApp.Documents.Add();

            wordDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;


            Word.Section section = wordDoc.Sections[1];
            section.PageSetup.TopMargin = wordApp.CentimetersToPoints(1.6f);
            section.PageSetup.LeftMargin = wordApp.CentimetersToPoints(1.0f);


            Word.Paragraph texxt = wordDoc.Content.Paragraphs.Add();
            texxt.Range.Text = "Отзыв руководителя от учреждения образования об учебной ";
            texxt.Range.Font.Size = 10;
            texxt.Range.Font.Name = "Times New Roman";
            texxt.Range.Font.Bold = 0;
            texxt.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            texxt.Range.ParagraphFormat.LeftIndent = wordApp.CentimetersToPoints(1.5f);
            texxt.Range.ParagraphFormat.SpaceAfter = 0;
            texxt.Range.InsertParagraphAfter();


            Word.Paragraph texxt1 = wordDoc.Content.Paragraphs.Add();
            texxt1.Range.Text = "практики";
            texxt1.Range.Font.Size = 10;
            texxt1.Range.Font.Name = "Times New Roman";
            texxt1.Range.Font.Bold = 0;
            texxt1.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            texxt1.Range.ParagraphFormat.LeftIndent = wordApp.CentimetersToPoints(1.5f);
            texxt1.Range.ParagraphFormat.SpaceAfter = 78;
            texxt1.Range.InsertParagraphAfter();

            Word.Paragraph texxt2 = wordDoc.Content.Paragraphs.Add();
            texxt2.Range.Text = "Отметка по учебной";
            texxt2.Range.Font.Size = 10;
            texxt2.Range.Font.Name = "Times New Roman";
            texxt2.Range.Font.Bold = 0;
            texxt2.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            texxt2.Range.ParagraphFormat.LeftIndent = wordApp.CentimetersToPoints(1.5f);
            texxt2.Range.ParagraphFormat.SpaceAfter = 0;
            texxt2.Range.InsertParagraphAfter();

            Word.Paragraph texxt3 = wordDoc.Content.Paragraphs.Add();
            texxt3.Range.Text = "практике";
            texxt3.Range.Font.Size = 10;
            texxt3.Range.Font.Name = "Times New Roman";
            texxt3.Range.Font.Bold = 0;
            texxt3.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            texxt3.Range.ParagraphFormat.LeftIndent = wordApp.CentimetersToPoints(1.5f);
            texxt3.Range.ParagraphFormat.SpaceAfter = 30;
            texxt3.Range.InsertParagraphAfter();


            Word.Paragraph texxt4 = wordDoc.Content.Paragraphs.Add();
            texxt4.Range.Text = "Подпись";
            texxt4.Range.Font.Size = 10;
            texxt4.Range.Font.Name = "Times New Roman";
            texxt4.Range.Font.Bold = 0;
            texxt4.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            texxt4.Range.ParagraphFormat.LeftIndent = wordApp.CentimetersToPoints(1.5f);
            texxt4.Range.ParagraphFormat.SpaceAfter = 30;
            texxt4.Range.InsertParagraphAfter();



            Word.Paragraph texxt5 = wordDoc.Content.Paragraphs.Add();
            texxt5.Range.Text = "«         »                    202    г.";
            texxt5.Range.Font.Size = 10;
            texxt5.Range.Font.Name = "Times New Roman";
            texxt5.Range.Font.Bold = 0;
            texxt5.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            texxt5.Range.ParagraphFormat.LeftIndent = wordApp.CentimetersToPoints(1.5f);
            texxt5.Range.ParagraphFormat.SpaceAfter = 30;
            texxt5.Range.InsertParagraphAfter();



            float lineWidth11 = 3.6f;
            Shape lineShape11 = wordDoc.Shapes.AddLine(
                wordApp.InchesToPoints(1.0f),
                wordApp.InchesToPoints(1.2f),
                wordApp.InchesToPoints(lineWidth11 + 1.0f),
                wordApp.InchesToPoints(1.2f)
            );


            lineShape11.Line.ForeColor.RGB = 0x000000;


            float lineWidth12 = 3.6f;
            Shape lineShape12 = wordDoc.Shapes.AddLine(
                wordApp.InchesToPoints(1.0f),
                wordApp.InchesToPoints(1.45f),
                wordApp.InchesToPoints(lineWidth12 + 1.0f),
                wordApp.InchesToPoints(1.45f)
            );


            lineShape12.Line.ForeColor.RGB = 0x000000;

            float lineWidth13 = 3.6f;
            Shape lineShape13 = wordDoc.Shapes.AddLine(
                wordApp.InchesToPoints(1.0f),
                wordApp.InchesToPoints(1.7f),
                wordApp.InchesToPoints(lineWidth13 + 1.0f),
                wordApp.InchesToPoints(1.7f)
            );


            lineShape13.Line.ForeColor.RGB = 0x000000;


            float lineWidth14 = 3.05f;
            Shape lineShape14 = wordDoc.Shapes.AddLine(
                wordApp.InchesToPoints(1.55f),
                wordApp.InchesToPoints(2.36f),
                wordApp.InchesToPoints(lineWidth14 + 1.55f),
                wordApp.InchesToPoints(2.36f)
            );


            lineShape14.Line.ForeColor.RGB = 0x000000;



            float lineWidth15 = 3.05f;
            Shape lineShape15 = wordDoc.Shapes.AddLine(
                wordApp.InchesToPoints(1.55f),
                wordApp.InchesToPoints(2.95f),
                wordApp.InchesToPoints(lineWidth15 + 1.55f),
                wordApp.InchesToPoints(2.95f)
            );


            lineShape15.Line.ForeColor.RGB = 0x000000;

            float lineWidth16 = 0.28f;
            Shape lineShape16 = wordDoc.Shapes.AddLine(
                wordApp.InchesToPoints(1.06f),
                wordApp.InchesToPoints(3.55f),
                wordApp.InchesToPoints(lineWidth16 + 1.06f),
                wordApp.InchesToPoints(3.55f)
            );


            lineShape16.Line.ForeColor.RGB = 0x000000;



            float lineWidth17 = 0.54f;
            Shape lineShape17 = wordDoc.Shapes.AddLine(
                wordApp.InchesToPoints(1.55f),
                wordApp.InchesToPoints(3.55f),
                wordApp.InchesToPoints(lineWidth17 + 1.55f),
                wordApp.InchesToPoints(3.55f)
            );


            lineShape17.Line.ForeColor.RGB = 0x000000;


            float lineWidth18 = 0.1f;
            Shape lineShape18 = wordDoc.Shapes.AddLine(
                wordApp.InchesToPoints(2.35f),
                wordApp.InchesToPoints(3.55f),
                wordApp.InchesToPoints(lineWidth18 + 2.35f),
                wordApp.InchesToPoints(3.55f)
            );


            lineShape18.Line.ForeColor.RGB = 0x000000;



            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "Документ Word (*.docx)|*.docx";
            saveFileDialog1.Title = "Сохранить документ Word";
            saveFileDialog1.ShowDialog();

            if (saveFileDialog1.FileName != "")
            {
                wordDoc.SaveAs2(saveFileDialog1.FileName);
            }


            wordDoc.Close();
            wordApp.Quit();

        }


    }
}
