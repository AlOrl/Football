using DocumentFormat.OpenXml.Packaging;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using M = DocumentFormat.OpenXml.Math;

namespace GeneratedCode
{
    public class GeneratedClass
    {
        // Creates a WordprocessingDocument.
        public void CreatePackage(string filePath)
        {
            using(WordprocessingDocument package = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
            {
                CreateParts(package);
            }
        }

        // Adds child parts and generates content of the specified part.
        private void CreateParts(WordprocessingDocument document)
        {
            ExtendedFilePropertiesPart extendedFilePropertiesPart1 = document.AddNewPart<ExtendedFilePropertiesPart>("rId3");
            GenerateExtendedFilePropertiesPart1Content(extendedFilePropertiesPart1);

            MainDocumentPart mainDocumentPart1 = document.AddMainDocumentPart();
            GenerateMainDocumentPart1Content(mainDocumentPart1);

            ThemePart themePart1 = mainDocumentPart1.AddNewPart<ThemePart>("rId8");
            GenerateThemePart1Content(themePart1);

            DocumentSettingsPart documentSettingsPart1 = mainDocumentPart1.AddNewPart<DocumentSettingsPart>("rId3");
            GenerateDocumentSettingsPart1Content(documentSettingsPart1);

            FontTablePart fontTablePart1 = mainDocumentPart1.AddNewPart<FontTablePart>("rId7");
            GenerateFontTablePart1Content(fontTablePart1);

            StyleDefinitionsPart styleDefinitionsPart1 = mainDocumentPart1.AddNewPart<StyleDefinitionsPart>("rId2");
            GenerateStyleDefinitionsPart1Content(styleDefinitionsPart1);

            NumberingDefinitionsPart numberingDefinitionsPart1 = mainDocumentPart1.AddNewPart<NumberingDefinitionsPart>("rId1");
            GenerateNumberingDefinitionsPart1Content(numberingDefinitionsPart1);

            EndnotesPart endnotesPart1 = mainDocumentPart1.AddNewPart<EndnotesPart>("rId6");
            GenerateEndnotesPart1Content(endnotesPart1);

            FootnotesPart footnotesPart1 = mainDocumentPart1.AddNewPart<FootnotesPart>("rId5");
            GenerateFootnotesPart1Content(footnotesPart1);

            WebSettingsPart webSettingsPart1 = mainDocumentPart1.AddNewPart<WebSettingsPart>("rId4");
            GenerateWebSettingsPart1Content(webSettingsPart1);

            SetPackageProperties(document);
        }

        // Generates content of extendedFilePropertiesPart1.
        private void GenerateExtendedFilePropertiesPart1Content(ExtendedFilePropertiesPart extendedFilePropertiesPart1)
        {
            Ap.Properties properties1 = new Ap.Properties();
            properties1.AddNamespaceDeclaration("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
            Ap.Template template1 = new Ap.Template();
            template1.Text = "Normal_Wordconv.dotm";
            Ap.TotalTime totalTime1 = new Ap.TotalTime();
            totalTime1.Text = "1";
            Ap.Pages pages1 = new Ap.Pages();
            pages1.Text = "1";
            Ap.Words words1 = new Ap.Words();
            words1.Text = "1";
            Ap.Characters characters1 = new Ap.Characters();
            characters1.Text = "9";
            Ap.Application application1 = new Ap.Application();
            application1.Text = "Microsoft Office Outlook";
            Ap.DocumentSecurity documentSecurity1 = new Ap.DocumentSecurity();
            documentSecurity1.Text = "0";
            Ap.Lines lines1 = new Ap.Lines();
            lines1.Text = "0";
            Ap.Paragraphs paragraphs1 = new Ap.Paragraphs();
            paragraphs1.Text = "0";
            Ap.ScaleCrop scaleCrop1 = new Ap.ScaleCrop();
            scaleCrop1.Text = "false";
            Ap.Company company1 = new Ap.Company();
            company1.Text = "";
            Ap.LinksUpToDate linksUpToDate1 = new Ap.LinksUpToDate();
            linksUpToDate1.Text = "false";
            Ap.CharactersWithSpaces charactersWithSpaces1 = new Ap.CharactersWithSpaces();
            charactersWithSpaces1.Text = "0";
            Ap.SharedDocument sharedDocument1 = new Ap.SharedDocument();
            sharedDocument1.Text = "false";
            Ap.HyperlinksChanged hyperlinksChanged1 = new Ap.HyperlinksChanged();
            hyperlinksChanged1.Text = "false";
            Ap.ApplicationVersion applicationVersion1 = new Ap.ApplicationVersion();
            applicationVersion1.Text = "12.0000";

            properties1.Append(template1);
            properties1.Append(totalTime1);
            properties1.Append(pages1);
            properties1.Append(words1);
            properties1.Append(characters1);
            properties1.Append(application1);
            properties1.Append(documentSecurity1);
            properties1.Append(lines1);
            properties1.Append(paragraphs1);
            properties1.Append(scaleCrop1);
            properties1.Append(company1);
            properties1.Append(linksUpToDate1);
            properties1.Append(charactersWithSpaces1);
            properties1.Append(sharedDocument1);
            properties1.Append(hyperlinksChanged1);
            properties1.Append(applicationVersion1);

            extendedFilePropertiesPart1.Properties = properties1;
        }

        // Generates content of mainDocumentPart1.
        private void GenerateMainDocumentPart1Content(MainDocumentPart mainDocumentPart1)
        {
            Document document1 = new Document();
            document1.AddNamespaceDeclaration("ve", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            document1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            document1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            document1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            document1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            document1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            document1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            document1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            document1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");

            Body body1 = new Body();

            Paragraph paragraph1 = new Paragraph(){ RsidParagraphMarkRevision = "006010A8", RsidParagraphAddition = "00C3696F", RsidParagraphProperties = "009C0664", RsidRunAdditionDefault = "00C3696F" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines(){ After = "0" };
            Justification justification1 = new Justification(){ Val = JustificationValues.Left };
            TextAlignment textAlignment1 = new TextAlignment(){ Val = VerticalTextAlignmentValues.Baseline };

            ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
            NoProof noProof1 = new NoProof();
            Languages languages1 = new Languages(){ Val = "en-US" };

            paragraphMarkRunProperties1.Append(noProof1);
            paragraphMarkRunProperties1.Append(languages1);

            paragraphProperties1.Append(spacingBetweenLines1);
            paragraphProperties1.Append(justification1);
            paragraphProperties1.Append(textAlignment1);
            paragraphProperties1.Append(paragraphMarkRunProperties1);

            Run run1 = new Run();

            RunProperties runProperties1 = new RunProperties();
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript(){ Val = "28" };

            

            runProperties1.Append(fontSizeComplexScript1);
            Text text1 = new Text();
            text1.Text = "jopa";

            run1.Append(runProperties1);
            run1.Append(text1);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);

            SectionProperties sectionProperties1 = new SectionProperties(){ RsidRPr = "006010A8", RsidR = "00C3696F", RsidSect = "00AA1605" };
            PageSize pageSize1 = new PageSize(){ Width = (UInt32Value)11906U, Height = (UInt32Value)16838U };
            PageMargin pageMargin1 = new PageMargin(){ Top = 284, Right = (UInt32Value)850U, Bottom = 1134, Left = (UInt32Value)1701U, Header = (UInt32Value)708U, Footer = (UInt32Value)708U, Gutter = (UInt32Value)0U };
            Columns columns1 = new Columns(){ Space = "708" };
            DocGrid docGrid1 = new DocGrid(){ LinePitch = 360 };

            sectionProperties1.Append(pageSize1);
            sectionProperties1.Append(pageMargin1);
            sectionProperties1.Append(columns1);
            sectionProperties1.Append(docGrid1);

            body1.Append(paragraph1);
            body1.Append(sectionProperties1);

            document1.Append(body1);

            mainDocumentPart1.Document = document1;
        }

        // Generates content of themePart1.
        private void GenerateThemePart1Content(ThemePart themePart1)
        {
            A.Theme theme1 = new A.Theme(){ Name = "Office Theme" };
            theme1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.ThemeElements themeElements1 = new A.ThemeElements();

            A.ColorScheme colorScheme1 = new A.ColorScheme(){ Name = "Office" };

            A.Dark1Color dark1Color1 = new A.Dark1Color();
            A.SystemColor systemColor1 = new A.SystemColor(){ Val = A.SystemColorValues.WindowText, LastColor = "000000" };

            dark1Color1.Append(systemColor1);

            A.Light1Color light1Color1 = new A.Light1Color();
            A.SystemColor systemColor2 = new A.SystemColor(){ Val = A.SystemColorValues.Window, LastColor = "FFFFFF" };

            light1Color1.Append(systemColor2);

            A.Dark2Color dark2Color1 = new A.Dark2Color();
            A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex(){ Val = "1F497D" };

            dark2Color1.Append(rgbColorModelHex1);

            A.Light2Color light2Color1 = new A.Light2Color();
            A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex(){ Val = "EEECE1" };

            light2Color1.Append(rgbColorModelHex2);

            A.Accent1Color accent1Color1 = new A.Accent1Color();
            A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex(){ Val = "4F81BD" };

            accent1Color1.Append(rgbColorModelHex3);

            A.Accent2Color accent2Color1 = new A.Accent2Color();
            A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex(){ Val = "C0504D" };

            accent2Color1.Append(rgbColorModelHex4);

            A.Accent3Color accent3Color1 = new A.Accent3Color();
            A.RgbColorModelHex rgbColorModelHex5 = new A.RgbColorModelHex(){ Val = "9BBB59" };

            accent3Color1.Append(rgbColorModelHex5);

            A.Accent4Color accent4Color1 = new A.Accent4Color();
            A.RgbColorModelHex rgbColorModelHex6 = new A.RgbColorModelHex(){ Val = "8064A2" };

            accent4Color1.Append(rgbColorModelHex6);

            A.Accent5Color accent5Color1 = new A.Accent5Color();
            A.RgbColorModelHex rgbColorModelHex7 = new A.RgbColorModelHex(){ Val = "4BACC6" };

            accent5Color1.Append(rgbColorModelHex7);

            A.Accent6Color accent6Color1 = new A.Accent6Color();
            A.RgbColorModelHex rgbColorModelHex8 = new A.RgbColorModelHex(){ Val = "F79646" };

            accent6Color1.Append(rgbColorModelHex8);

            A.Hyperlink hyperlink1 = new A.Hyperlink();
            A.RgbColorModelHex rgbColorModelHex9 = new A.RgbColorModelHex(){ Val = "0000FF" };

            hyperlink1.Append(rgbColorModelHex9);

            A.FollowedHyperlinkColor followedHyperlinkColor1 = new A.FollowedHyperlinkColor();
            A.RgbColorModelHex rgbColorModelHex10 = new A.RgbColorModelHex(){ Val = "800080" };

            followedHyperlinkColor1.Append(rgbColorModelHex10);

            colorScheme1.Append(dark1Color1);
            colorScheme1.Append(light1Color1);
            colorScheme1.Append(dark2Color1);
            colorScheme1.Append(light2Color1);
            colorScheme1.Append(accent1Color1);
            colorScheme1.Append(accent2Color1);
            colorScheme1.Append(accent3Color1);
            colorScheme1.Append(accent4Color1);
            colorScheme1.Append(accent5Color1);
            colorScheme1.Append(accent6Color1);
            colorScheme1.Append(hyperlink1);
            colorScheme1.Append(followedHyperlinkColor1);

            A.FontScheme fontScheme1 = new A.FontScheme(){ Name = "Office" };

            A.MajorFont majorFont1 = new A.MajorFont();
            A.LatinFont latinFont1 = new A.LatinFont(){ Typeface = "Cambria" };
            A.EastAsianFont eastAsianFont1 = new A.EastAsianFont(){ Typeface = "" };
            A.ComplexScriptFont complexScriptFont1 = new A.ComplexScriptFont(){ Typeface = "" };
            A.SupplementalFont supplementalFont1 = new A.SupplementalFont(){ Script = "Jpan", Typeface = "ＭＳ ゴシック" };
            A.SupplementalFont supplementalFont2 = new A.SupplementalFont(){ Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont3 = new A.SupplementalFont(){ Script = "Hans", Typeface = "宋体" };
            A.SupplementalFont supplementalFont4 = new A.SupplementalFont(){ Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont5 = new A.SupplementalFont(){ Script = "Arab", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont6 = new A.SupplementalFont(){ Script = "Hebr", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont7 = new A.SupplementalFont(){ Script = "Thai", Typeface = "Angsana New" };
            A.SupplementalFont supplementalFont8 = new A.SupplementalFont(){ Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont9 = new A.SupplementalFont(){ Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont10 = new A.SupplementalFont(){ Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont11 = new A.SupplementalFont(){ Script = "Khmr", Typeface = "MoolBoran" };
            A.SupplementalFont supplementalFont12 = new A.SupplementalFont(){ Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont13 = new A.SupplementalFont(){ Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont14 = new A.SupplementalFont(){ Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont15 = new A.SupplementalFont(){ Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont16 = new A.SupplementalFont(){ Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont17 = new A.SupplementalFont(){ Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont18 = new A.SupplementalFont(){ Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont19 = new A.SupplementalFont(){ Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont20 = new A.SupplementalFont(){ Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont21 = new A.SupplementalFont(){ Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont22 = new A.SupplementalFont(){ Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont23 = new A.SupplementalFont(){ Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont24 = new A.SupplementalFont(){ Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont25 = new A.SupplementalFont(){ Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont26 = new A.SupplementalFont(){ Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont27 = new A.SupplementalFont(){ Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont28 = new A.SupplementalFont(){ Script = "Viet", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont29 = new A.SupplementalFont(){ Script = "Uigh", Typeface = "Microsoft Uighur" };

            majorFont1.Append(latinFont1);
            majorFont1.Append(eastAsianFont1);
            majorFont1.Append(complexScriptFont1);
            majorFont1.Append(supplementalFont1);
            majorFont1.Append(supplementalFont2);
            majorFont1.Append(supplementalFont3);
            majorFont1.Append(supplementalFont4);
            majorFont1.Append(supplementalFont5);
            majorFont1.Append(supplementalFont6);
            majorFont1.Append(supplementalFont7);
            majorFont1.Append(supplementalFont8);
            majorFont1.Append(supplementalFont9);
            majorFont1.Append(supplementalFont10);
            majorFont1.Append(supplementalFont11);
            majorFont1.Append(supplementalFont12);
            majorFont1.Append(supplementalFont13);
            majorFont1.Append(supplementalFont14);
            majorFont1.Append(supplementalFont15);
            majorFont1.Append(supplementalFont16);
            majorFont1.Append(supplementalFont17);
            majorFont1.Append(supplementalFont18);
            majorFont1.Append(supplementalFont19);
            majorFont1.Append(supplementalFont20);
            majorFont1.Append(supplementalFont21);
            majorFont1.Append(supplementalFont22);
            majorFont1.Append(supplementalFont23);
            majorFont1.Append(supplementalFont24);
            majorFont1.Append(supplementalFont25);
            majorFont1.Append(supplementalFont26);
            majorFont1.Append(supplementalFont27);
            majorFont1.Append(supplementalFont28);
            majorFont1.Append(supplementalFont29);

            A.MinorFont minorFont1 = new A.MinorFont();
            A.LatinFont latinFont2 = new A.LatinFont(){ Typeface = "Calibri" };
            A.EastAsianFont eastAsianFont2 = new A.EastAsianFont(){ Typeface = "" };
            A.ComplexScriptFont complexScriptFont2 = new A.ComplexScriptFont(){ Typeface = "" };
            A.SupplementalFont supplementalFont30 = new A.SupplementalFont(){ Script = "Jpan", Typeface = "ＭＳ 明朝" };
            A.SupplementalFont supplementalFont31 = new A.SupplementalFont(){ Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont32 = new A.SupplementalFont(){ Script = "Hans", Typeface = "宋体" };
            A.SupplementalFont supplementalFont33 = new A.SupplementalFont(){ Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont34 = new A.SupplementalFont(){ Script = "Arab", Typeface = "Arial" };
            A.SupplementalFont supplementalFont35 = new A.SupplementalFont(){ Script = "Hebr", Typeface = "Arial" };
            A.SupplementalFont supplementalFont36 = new A.SupplementalFont(){ Script = "Thai", Typeface = "Cordia New" };
            A.SupplementalFont supplementalFont37 = new A.SupplementalFont(){ Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont38 = new A.SupplementalFont(){ Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont39 = new A.SupplementalFont(){ Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont40 = new A.SupplementalFont(){ Script = "Khmr", Typeface = "DaunPenh" };
            A.SupplementalFont supplementalFont41 = new A.SupplementalFont(){ Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont42 = new A.SupplementalFont(){ Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont43 = new A.SupplementalFont(){ Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont44 = new A.SupplementalFont(){ Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont45 = new A.SupplementalFont(){ Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont46 = new A.SupplementalFont(){ Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont47 = new A.SupplementalFont(){ Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont48 = new A.SupplementalFont(){ Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont49 = new A.SupplementalFont(){ Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont50 = new A.SupplementalFont(){ Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont51 = new A.SupplementalFont(){ Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont52 = new A.SupplementalFont(){ Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont53 = new A.SupplementalFont(){ Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont54 = new A.SupplementalFont(){ Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont55 = new A.SupplementalFont(){ Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont56 = new A.SupplementalFont(){ Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont57 = new A.SupplementalFont(){ Script = "Viet", Typeface = "Arial" };
            A.SupplementalFont supplementalFont58 = new A.SupplementalFont(){ Script = "Uigh", Typeface = "Microsoft Uighur" };

            minorFont1.Append(latinFont2);
            minorFont1.Append(eastAsianFont2);
            minorFont1.Append(complexScriptFont2);
            minorFont1.Append(supplementalFont30);
            minorFont1.Append(supplementalFont31);
            minorFont1.Append(supplementalFont32);
            minorFont1.Append(supplementalFont33);
            minorFont1.Append(supplementalFont34);
            minorFont1.Append(supplementalFont35);
            minorFont1.Append(supplementalFont36);
            minorFont1.Append(supplementalFont37);
            minorFont1.Append(supplementalFont38);
            minorFont1.Append(supplementalFont39);
            minorFont1.Append(supplementalFont40);
            minorFont1.Append(supplementalFont41);
            minorFont1.Append(supplementalFont42);
            minorFont1.Append(supplementalFont43);
            minorFont1.Append(supplementalFont44);
            minorFont1.Append(supplementalFont45);
            minorFont1.Append(supplementalFont46);
            minorFont1.Append(supplementalFont47);
            minorFont1.Append(supplementalFont48);
            minorFont1.Append(supplementalFont49);
            minorFont1.Append(supplementalFont50);
            minorFont1.Append(supplementalFont51);
            minorFont1.Append(supplementalFont52);
            minorFont1.Append(supplementalFont53);
            minorFont1.Append(supplementalFont54);
            minorFont1.Append(supplementalFont55);
            minorFont1.Append(supplementalFont56);
            minorFont1.Append(supplementalFont57);
            minorFont1.Append(supplementalFont58);

            fontScheme1.Append(majorFont1);
            fontScheme1.Append(minorFont1);

            A.FormatScheme formatScheme1 = new A.FormatScheme(){ Name = "Office" };

            A.FillStyleList fillStyleList1 = new A.FillStyleList();

            A.SolidFill solidFill1 = new A.SolidFill();
            A.SchemeColor schemeColor1 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };

            solidFill1.Append(schemeColor1);

            A.GradientFill gradientFill1 = new A.GradientFill(){ RotateWithShape = true };

            A.GradientStopList gradientStopList1 = new A.GradientStopList();

            A.GradientStop gradientStop1 = new A.GradientStop(){ Position = 0 };

            A.SchemeColor schemeColor2 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.Tint tint1 = new A.Tint(){ Val = 50000 };
            A.SaturationModulation saturationModulation1 = new A.SaturationModulation(){ Val = 300000 };

            schemeColor2.Append(tint1);
            schemeColor2.Append(saturationModulation1);

            gradientStop1.Append(schemeColor2);

            A.GradientStop gradientStop2 = new A.GradientStop(){ Position = 35000 };

            A.SchemeColor schemeColor3 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.Tint tint2 = new A.Tint(){ Val = 37000 };
            A.SaturationModulation saturationModulation2 = new A.SaturationModulation(){ Val = 300000 };

            schemeColor3.Append(tint2);
            schemeColor3.Append(saturationModulation2);

            gradientStop2.Append(schemeColor3);

            A.GradientStop gradientStop3 = new A.GradientStop(){ Position = 100000 };

            A.SchemeColor schemeColor4 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.Tint tint3 = new A.Tint(){ Val = 15000 };
            A.SaturationModulation saturationModulation3 = new A.SaturationModulation(){ Val = 350000 };

            schemeColor4.Append(tint3);
            schemeColor4.Append(saturationModulation3);

            gradientStop3.Append(schemeColor4);

            gradientStopList1.Append(gradientStop1);
            gradientStopList1.Append(gradientStop2);
            gradientStopList1.Append(gradientStop3);
            A.LinearGradientFill linearGradientFill1 = new A.LinearGradientFill(){ Angle = 16200000, Scaled = true };

            gradientFill1.Append(gradientStopList1);
            gradientFill1.Append(linearGradientFill1);

            A.GradientFill gradientFill2 = new A.GradientFill(){ RotateWithShape = true };

            A.GradientStopList gradientStopList2 = new A.GradientStopList();

            A.GradientStop gradientStop4 = new A.GradientStop(){ Position = 0 };

            A.SchemeColor schemeColor5 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.Shade shade1 = new A.Shade(){ Val = 51000 };
            A.SaturationModulation saturationModulation4 = new A.SaturationModulation(){ Val = 130000 };

            schemeColor5.Append(shade1);
            schemeColor5.Append(saturationModulation4);

            gradientStop4.Append(schemeColor5);

            A.GradientStop gradientStop5 = new A.GradientStop(){ Position = 80000 };

            A.SchemeColor schemeColor6 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.Shade shade2 = new A.Shade(){ Val = 93000 };
            A.SaturationModulation saturationModulation5 = new A.SaturationModulation(){ Val = 130000 };

            schemeColor6.Append(shade2);
            schemeColor6.Append(saturationModulation5);

            gradientStop5.Append(schemeColor6);

            A.GradientStop gradientStop6 = new A.GradientStop(){ Position = 100000 };

            A.SchemeColor schemeColor7 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.Shade shade3 = new A.Shade(){ Val = 94000 };
            A.SaturationModulation saturationModulation6 = new A.SaturationModulation(){ Val = 135000 };

            schemeColor7.Append(shade3);
            schemeColor7.Append(saturationModulation6);

            gradientStop6.Append(schemeColor7);

            gradientStopList2.Append(gradientStop4);
            gradientStopList2.Append(gradientStop5);
            gradientStopList2.Append(gradientStop6);
            A.LinearGradientFill linearGradientFill2 = new A.LinearGradientFill(){ Angle = 16200000, Scaled = false };

            gradientFill2.Append(gradientStopList2);
            gradientFill2.Append(linearGradientFill2);

            fillStyleList1.Append(solidFill1);
            fillStyleList1.Append(gradientFill1);
            fillStyleList1.Append(gradientFill2);

            A.LineStyleList lineStyleList1 = new A.LineStyleList();

            A.Outline outline1 = new A.Outline(){ Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill2 = new A.SolidFill();

            A.SchemeColor schemeColor8 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.Shade shade4 = new A.Shade(){ Val = 95000 };
            A.SaturationModulation saturationModulation7 = new A.SaturationModulation(){ Val = 105000 };

            schemeColor8.Append(shade4);
            schemeColor8.Append(saturationModulation7);

            solidFill2.Append(schemeColor8);
            A.PresetDash presetDash1 = new A.PresetDash(){ Val = A.PresetLineDashValues.Solid };

            outline1.Append(solidFill2);
            outline1.Append(presetDash1);

            A.Outline outline2 = new A.Outline(){ Width = 25400, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill3 = new A.SolidFill();
            A.SchemeColor schemeColor9 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };

            solidFill3.Append(schemeColor9);
            A.PresetDash presetDash2 = new A.PresetDash(){ Val = A.PresetLineDashValues.Solid };

            outline2.Append(solidFill3);
            outline2.Append(presetDash2);

            A.Outline outline3 = new A.Outline(){ Width = 38100, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill4 = new A.SolidFill();
            A.SchemeColor schemeColor10 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };

            solidFill4.Append(schemeColor10);
            A.PresetDash presetDash3 = new A.PresetDash(){ Val = A.PresetLineDashValues.Solid };

            outline3.Append(solidFill4);
            outline3.Append(presetDash3);

            lineStyleList1.Append(outline1);
            lineStyleList1.Append(outline2);
            lineStyleList1.Append(outline3);

            A.EffectStyleList effectStyleList1 = new A.EffectStyleList();

            A.EffectStyle effectStyle1 = new A.EffectStyle();

            A.EffectList effectList1 = new A.EffectList();

            A.OuterShadow outerShadow1 = new A.OuterShadow(){ BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex11 = new A.RgbColorModelHex(){ Val = "000000" };
            A.Alpha alpha1 = new A.Alpha(){ Val = 38000 };

            rgbColorModelHex11.Append(alpha1);

            outerShadow1.Append(rgbColorModelHex11);

            effectList1.Append(outerShadow1);

            effectStyle1.Append(effectList1);

            A.EffectStyle effectStyle2 = new A.EffectStyle();

            A.EffectList effectList2 = new A.EffectList();

            A.OuterShadow outerShadow2 = new A.OuterShadow(){ BlurRadius = 40000L, Distance = 23000L, Direction = 5400000, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex12 = new A.RgbColorModelHex(){ Val = "000000" };
            A.Alpha alpha2 = new A.Alpha(){ Val = 35000 };

            rgbColorModelHex12.Append(alpha2);

            outerShadow2.Append(rgbColorModelHex12);

            effectList2.Append(outerShadow2);

            effectStyle2.Append(effectList2);

            A.EffectStyle effectStyle3 = new A.EffectStyle();

            A.EffectList effectList3 = new A.EffectList();

            A.OuterShadow outerShadow3 = new A.OuterShadow(){ BlurRadius = 40000L, Distance = 23000L, Direction = 5400000, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex13 = new A.RgbColorModelHex(){ Val = "000000" };
            A.Alpha alpha3 = new A.Alpha(){ Val = 35000 };

            rgbColorModelHex13.Append(alpha3);

            outerShadow3.Append(rgbColorModelHex13);

            effectList3.Append(outerShadow3);

            A.Scene3DType scene3DType1 = new A.Scene3DType();

            A.Camera camera1 = new A.Camera(){ Preset = A.PresetCameraValues.OrthographicFront };
            A.Rotation rotation1 = new A.Rotation(){ Latitude = 0, Longitude = 0, Revolution = 0 };

            camera1.Append(rotation1);

            A.LightRig lightRig1 = new A.LightRig(){ Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };
            A.Rotation rotation2 = new A.Rotation(){ Latitude = 0, Longitude = 0, Revolution = 1200000 };

            lightRig1.Append(rotation2);

            scene3DType1.Append(camera1);
            scene3DType1.Append(lightRig1);

            A.Shape3DType shape3DType1 = new A.Shape3DType();
            A.BevelTop bevelTop1 = new A.BevelTop(){ Width = 63500L, Height = 25400L };

            shape3DType1.Append(bevelTop1);

            effectStyle3.Append(effectList3);
            effectStyle3.Append(scene3DType1);
            effectStyle3.Append(shape3DType1);

            effectStyleList1.Append(effectStyle1);
            effectStyleList1.Append(effectStyle2);
            effectStyleList1.Append(effectStyle3);

            A.BackgroundFillStyleList backgroundFillStyleList1 = new A.BackgroundFillStyleList();

            A.SolidFill solidFill5 = new A.SolidFill();
            A.SchemeColor schemeColor11 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };

            solidFill5.Append(schemeColor11);

            A.GradientFill gradientFill3 = new A.GradientFill(){ RotateWithShape = true };

            A.GradientStopList gradientStopList3 = new A.GradientStopList();

            A.GradientStop gradientStop7 = new A.GradientStop(){ Position = 0 };

            A.SchemeColor schemeColor12 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.Tint tint4 = new A.Tint(){ Val = 40000 };
            A.SaturationModulation saturationModulation8 = new A.SaturationModulation(){ Val = 350000 };

            schemeColor12.Append(tint4);
            schemeColor12.Append(saturationModulation8);

            gradientStop7.Append(schemeColor12);

            A.GradientStop gradientStop8 = new A.GradientStop(){ Position = 40000 };

            A.SchemeColor schemeColor13 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.Tint tint5 = new A.Tint(){ Val = 45000 };
            A.Shade shade5 = new A.Shade(){ Val = 99000 };
            A.SaturationModulation saturationModulation9 = new A.SaturationModulation(){ Val = 350000 };

            schemeColor13.Append(tint5);
            schemeColor13.Append(shade5);
            schemeColor13.Append(saturationModulation9);

            gradientStop8.Append(schemeColor13);

            A.GradientStop gradientStop9 = new A.GradientStop(){ Position = 100000 };

            A.SchemeColor schemeColor14 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.Shade shade6 = new A.Shade(){ Val = 20000 };
            A.SaturationModulation saturationModulation10 = new A.SaturationModulation(){ Val = 255000 };

            schemeColor14.Append(shade6);
            schemeColor14.Append(saturationModulation10);

            gradientStop9.Append(schemeColor14);

            gradientStopList3.Append(gradientStop7);
            gradientStopList3.Append(gradientStop8);
            gradientStopList3.Append(gradientStop9);

            A.PathGradientFill pathGradientFill1 = new A.PathGradientFill(){ Path = A.PathShadeValues.Circle };
            A.FillToRectangle fillToRectangle1 = new A.FillToRectangle(){ Left = 50000, Top = -80000, Right = 50000, Bottom = 180000 };

            pathGradientFill1.Append(fillToRectangle1);

            gradientFill3.Append(gradientStopList3);
            gradientFill3.Append(pathGradientFill1);

            A.GradientFill gradientFill4 = new A.GradientFill(){ RotateWithShape = true };

            A.GradientStopList gradientStopList4 = new A.GradientStopList();

            A.GradientStop gradientStop10 = new A.GradientStop(){ Position = 0 };

            A.SchemeColor schemeColor15 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.Tint tint6 = new A.Tint(){ Val = 80000 };
            A.SaturationModulation saturationModulation11 = new A.SaturationModulation(){ Val = 300000 };

            schemeColor15.Append(tint6);
            schemeColor15.Append(saturationModulation11);

            gradientStop10.Append(schemeColor15);

            A.GradientStop gradientStop11 = new A.GradientStop(){ Position = 100000 };

            A.SchemeColor schemeColor16 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.Shade shade7 = new A.Shade(){ Val = 30000 };
            A.SaturationModulation saturationModulation12 = new A.SaturationModulation(){ Val = 200000 };

            schemeColor16.Append(shade7);
            schemeColor16.Append(saturationModulation12);

            gradientStop11.Append(schemeColor16);

            gradientStopList4.Append(gradientStop10);
            gradientStopList4.Append(gradientStop11);

            A.PathGradientFill pathGradientFill2 = new A.PathGradientFill(){ Path = A.PathShadeValues.Circle };
            A.FillToRectangle fillToRectangle2 = new A.FillToRectangle(){ Left = 50000, Top = 50000, Right = 50000, Bottom = 50000 };

            pathGradientFill2.Append(fillToRectangle2);

            gradientFill4.Append(gradientStopList4);
            gradientFill4.Append(pathGradientFill2);

            backgroundFillStyleList1.Append(solidFill5);
            backgroundFillStyleList1.Append(gradientFill3);
            backgroundFillStyleList1.Append(gradientFill4);

            formatScheme1.Append(fillStyleList1);
            formatScheme1.Append(lineStyleList1);
            formatScheme1.Append(effectStyleList1);
            formatScheme1.Append(backgroundFillStyleList1);

            themeElements1.Append(colorScheme1);
            themeElements1.Append(fontScheme1);
            themeElements1.Append(formatScheme1);
            A.ObjectDefaults objectDefaults1 = new A.ObjectDefaults();
            A.ExtraColorSchemeList extraColorSchemeList1 = new A.ExtraColorSchemeList();

            theme1.Append(themeElements1);
            theme1.Append(objectDefaults1);
            theme1.Append(extraColorSchemeList1);

            themePart1.Theme = theme1;
        }

        // Generates content of documentSettingsPart1.
        private void GenerateDocumentSettingsPart1Content(DocumentSettingsPart documentSettingsPart1)
        {
            Settings settings1 = new Settings();
            settings1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            settings1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            settings1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            settings1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            settings1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            settings1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            settings1.AddNamespaceDeclaration("sl", "http://schemas.openxmlformats.org/schemaLibrary/2006/main");
            Zoom zoom1 = new Zoom(){ Percent = "100" };
            DefaultTabStop defaultTabStop1 = new DefaultTabStop(){ Val = 708 };
            CharacterSpacingControl characterSpacingControl1 = new CharacterSpacingControl(){ Val = CharacterSpacingValues.DoNotCompress };

            FootnoteDocumentWideProperties footnoteDocumentWideProperties1 = new FootnoteDocumentWideProperties();
            FootnoteSpecialReference footnoteSpecialReference1 = new FootnoteSpecialReference(){ Id = -1 };
            FootnoteSpecialReference footnoteSpecialReference2 = new FootnoteSpecialReference(){ Id = 0 };

            footnoteDocumentWideProperties1.Append(footnoteSpecialReference1);
            footnoteDocumentWideProperties1.Append(footnoteSpecialReference2);

            EndnoteDocumentWideProperties endnoteDocumentWideProperties1 = new EndnoteDocumentWideProperties();
            EndnoteSpecialReference endnoteSpecialReference1 = new EndnoteSpecialReference(){ Id = -1 };
            EndnoteSpecialReference endnoteSpecialReference2 = new EndnoteSpecialReference(){ Id = 0 };

            endnoteDocumentWideProperties1.Append(endnoteSpecialReference1);
            endnoteDocumentWideProperties1.Append(endnoteSpecialReference2);

            Compatibility compatibility1 = new Compatibility();
            UseNormalStyleForList useNormalStyleForList1 = new UseNormalStyleForList();
            DoNotUseIndentAsNumberingTabStop doNotUseIndentAsNumberingTabStop1 = new DoNotUseIndentAsNumberingTabStop();
            UseAltKinsokuLineBreakRules useAltKinsokuLineBreakRules1 = new UseAltKinsokuLineBreakRules();
            AllowSpaceOfSameStyleInTable allowSpaceOfSameStyleInTable1 = new AllowSpaceOfSameStyleInTable();
            DoNotSuppressIndentation doNotSuppressIndentation1 = new DoNotSuppressIndentation();
            DoNotAutofitConstrainedTables doNotAutofitConstrainedTables1 = new DoNotAutofitConstrainedTables();
            AutofitToFirstFixedWidthCell autofitToFirstFixedWidthCell1 = new AutofitToFirstFixedWidthCell();
            UnderlineTabInNumberingList underlineTabInNumberingList1 = new UnderlineTabInNumberingList();
            DisplayHangulFixedWidth displayHangulFixedWidth1 = new DisplayHangulFixedWidth();
            SplitPageBreakAndParagraphMark splitPageBreakAndParagraphMark1 = new SplitPageBreakAndParagraphMark();
            DoNotVerticallyAlignCellWithShape doNotVerticallyAlignCellWithShape1 = new DoNotVerticallyAlignCellWithShape();
            DoNotBreakConstrainedForcedTable doNotBreakConstrainedForcedTable1 = new DoNotBreakConstrainedForcedTable();
            DoNotVerticallyAlignInTextBox doNotVerticallyAlignInTextBox1 = new DoNotVerticallyAlignInTextBox();
            UseAnsiKerningPairs useAnsiKerningPairs1 = new UseAnsiKerningPairs();
            CachedColumnBalance cachedColumnBalance1 = new CachedColumnBalance();

            compatibility1.Append(useNormalStyleForList1);
            compatibility1.Append(doNotUseIndentAsNumberingTabStop1);
            compatibility1.Append(useAltKinsokuLineBreakRules1);
            compatibility1.Append(allowSpaceOfSameStyleInTable1);
            compatibility1.Append(doNotSuppressIndentation1);
            compatibility1.Append(doNotAutofitConstrainedTables1);
            compatibility1.Append(autofitToFirstFixedWidthCell1);
            compatibility1.Append(underlineTabInNumberingList1);
            compatibility1.Append(displayHangulFixedWidth1);
            compatibility1.Append(splitPageBreakAndParagraphMark1);
            compatibility1.Append(doNotVerticallyAlignCellWithShape1);
            compatibility1.Append(doNotBreakConstrainedForcedTable1);
            compatibility1.Append(doNotVerticallyAlignInTextBox1);
            compatibility1.Append(useAnsiKerningPairs1);
            compatibility1.Append(cachedColumnBalance1);

            Rsids rsids1 = new Rsids();
            RsidRoot rsidRoot1 = new RsidRoot(){ Val = "00CD6E14" };
            Rsid rsid1 = new Rsid(){ Val = "0005787C" };
            Rsid rsid2 = new Rsid(){ Val = "00071E5A" };
            Rsid rsid3 = new Rsid(){ Val = "00091877" };
            Rsid rsid4 = new Rsid(){ Val = "00107029" };
            Rsid rsid5 = new Rsid(){ Val = "00177A18" };
            Rsid rsid6 = new Rsid(){ Val = "001B4B2A" };
            Rsid rsid7 = new Rsid(){ Val = "00245160" };
            Rsid rsid8 = new Rsid(){ Val = "002E23DE" };
            Rsid rsid9 = new Rsid(){ Val = "00353D51" };
            Rsid rsid10 = new Rsid(){ Val = "003D1697" };
            Rsid rsid11 = new Rsid(){ Val = "00441DA6" };
            Rsid rsid12 = new Rsid(){ Val = "00444320" };
            Rsid rsid13 = new Rsid(){ Val = "00511120" };
            Rsid rsid14 = new Rsid(){ Val = "00552817" };
            Rsid rsid15 = new Rsid(){ Val = "006010A8" };
            Rsid rsid16 = new Rsid(){ Val = "00620D7D" };
            Rsid rsid17 = new Rsid(){ Val = "00624A44" };
            Rsid rsid18 = new Rsid(){ Val = "006E42F0" };
            Rsid rsid19 = new Rsid(){ Val = "006F04FA" };
            Rsid rsid20 = new Rsid(){ Val = "007A5696" };
            Rsid rsid21 = new Rsid(){ Val = "007D2F02" };
            Rsid rsid22 = new Rsid(){ Val = "007F2D1A" };
            Rsid rsid23 = new Rsid(){ Val = "008855B3" };
            Rsid rsid24 = new Rsid(){ Val = "00900C1F" };
            Rsid rsid25 = new Rsid(){ Val = "0094535F" };
            Rsid rsid26 = new Rsid(){ Val = "0098617D" };
            Rsid rsid27 = new Rsid(){ Val = "009C0664" };
            Rsid rsid28 = new Rsid(){ Val = "00AA1605" };
            Rsid rsid29 = new Rsid(){ Val = "00AC2DE8" };
            Rsid rsid30 = new Rsid(){ Val = "00BC03F4" };
            Rsid rsid31 = new Rsid(){ Val = "00C3696F" };
            Rsid rsid32 = new Rsid(){ Val = "00CD6E14" };
            Rsid rsid33 = new Rsid(){ Val = "00CE245C" };
            Rsid rsid34 = new Rsid(){ Val = "00D21418" };
            Rsid rsid35 = new Rsid(){ Val = "00D62D60" };
            Rsid rsid36 = new Rsid(){ Val = "00DA7E99" };
            Rsid rsid37 = new Rsid(){ Val = "00DC0C26" };
            Rsid rsid38 = new Rsid(){ Val = "00DD5B2E" };
            Rsid rsid39 = new Rsid(){ Val = "00DE6A58" };
            Rsid rsid40 = new Rsid(){ Val = "00E56771" };
            Rsid rsid41 = new Rsid(){ Val = "00E601AA" };
            Rsid rsid42 = new Rsid(){ Val = "00F10B91" };
            Rsid rsid43 = new Rsid(){ Val = "00F74B79" };
            Rsid rsid44 = new Rsid(){ Val = "00FA1958" };
            Rsid rsid45 = new Rsid(){ Val = "00FA7479" };

            rsids1.Append(rsidRoot1);
            rsids1.Append(rsid1);
            rsids1.Append(rsid2);
            rsids1.Append(rsid3);
            rsids1.Append(rsid4);
            rsids1.Append(rsid5);
            rsids1.Append(rsid6);
            rsids1.Append(rsid7);
            rsids1.Append(rsid8);
            rsids1.Append(rsid9);
            rsids1.Append(rsid10);
            rsids1.Append(rsid11);
            rsids1.Append(rsid12);
            rsids1.Append(rsid13);
            rsids1.Append(rsid14);
            rsids1.Append(rsid15);
            rsids1.Append(rsid16);
            rsids1.Append(rsid17);
            rsids1.Append(rsid18);
            rsids1.Append(rsid19);
            rsids1.Append(rsid20);
            rsids1.Append(rsid21);
            rsids1.Append(rsid22);
            rsids1.Append(rsid23);
            rsids1.Append(rsid24);
            rsids1.Append(rsid25);
            rsids1.Append(rsid26);
            rsids1.Append(rsid27);
            rsids1.Append(rsid28);
            rsids1.Append(rsid29);
            rsids1.Append(rsid30);
            rsids1.Append(rsid31);
            rsids1.Append(rsid32);
            rsids1.Append(rsid33);
            rsids1.Append(rsid34);
            rsids1.Append(rsid35);
            rsids1.Append(rsid36);
            rsids1.Append(rsid37);
            rsids1.Append(rsid38);
            rsids1.Append(rsid39);
            rsids1.Append(rsid40);
            rsids1.Append(rsid41);
            rsids1.Append(rsid42);
            rsids1.Append(rsid43);
            rsids1.Append(rsid44);
            rsids1.Append(rsid45);

            M.MathProperties mathProperties1 = new M.MathProperties();
            M.MathFont mathFont1 = new M.MathFont(){ Val = "Cambria Math" };
            M.BreakBinary breakBinary1 = new M.BreakBinary(){ Val = M.BreakBinaryOperatorValues.Before };
            M.BreakBinarySubtraction breakBinarySubtraction1 = new M.BreakBinarySubtraction(){ Val = M.BreakBinarySubtractionValues.MinusMinus };
            M.SmallFraction smallFraction1 = new M.SmallFraction(){ Val = M.BooleanValues.Off };
            M.DisplayDefaults displayDefaults1 = new M.DisplayDefaults();
            M.LeftMargin leftMargin1 = new M.LeftMargin(){ Val = (UInt32Value)0U };
            M.RightMargin rightMargin1 = new M.RightMargin(){ Val = (UInt32Value)0U };
            M.DefaultJustification defaultJustification1 = new M.DefaultJustification(){ Val = M.JustificationValues.CenterGroup };
            M.WrapIndent wrapIndent1 = new M.WrapIndent(){ Val = (UInt32Value)1440U };
            M.IntegralLimitLocation integralLimitLocation1 = new M.IntegralLimitLocation(){ Val = M.LimitLocationValues.SubscriptSuperscript };
            M.NaryLimitLocation naryLimitLocation1 = new M.NaryLimitLocation(){ Val = M.LimitLocationValues.UnderOver };

            mathProperties1.Append(mathFont1);
            mathProperties1.Append(breakBinary1);
            mathProperties1.Append(breakBinarySubtraction1);
            mathProperties1.Append(smallFraction1);
            mathProperties1.Append(displayDefaults1);
            mathProperties1.Append(leftMargin1);
            mathProperties1.Append(rightMargin1);
            mathProperties1.Append(defaultJustification1);
            mathProperties1.Append(wrapIndent1);
            mathProperties1.Append(integralLimitLocation1);
            mathProperties1.Append(naryLimitLocation1);
            UICompatibleWith97To2003 uICompatibleWith97To20031 = new UICompatibleWith97To2003();
            ThemeFontLanguages themeFontLanguages1 = new ThemeFontLanguages(){ Val = "ru-RU" };
            ColorSchemeMapping colorSchemeMapping1 = new ColorSchemeMapping(){ Background1 = ColorSchemeIndexValues.Light1, Text1 = ColorSchemeIndexValues.Dark1, Background2 = ColorSchemeIndexValues.Light2, Text2 = ColorSchemeIndexValues.Dark2, Accent1 = ColorSchemeIndexValues.Accent1, Accent2 = ColorSchemeIndexValues.Accent2, Accent3 = ColorSchemeIndexValues.Accent3, Accent4 = ColorSchemeIndexValues.Accent4, Accent5 = ColorSchemeIndexValues.Accent5, Accent6 = ColorSchemeIndexValues.Accent6, Hyperlink = ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = ColorSchemeIndexValues.FollowedHyperlink };
            DoNotIncludeSubdocsInStats doNotIncludeSubdocsInStats1 = new DoNotIncludeSubdocsInStats();
            DoNotAutoCompressPictures doNotAutoCompressPictures1 = new DoNotAutoCompressPictures();
            DecimalSymbol decimalSymbol1 = new DecimalSymbol(){ Val = "," };
            ListSeparator listSeparator1 = new ListSeparator(){ Val = ";" };

            settings1.Append(zoom1);
            settings1.Append(defaultTabStop1);
            settings1.Append(characterSpacingControl1);
            settings1.Append(footnoteDocumentWideProperties1);
            settings1.Append(endnoteDocumentWideProperties1);
            settings1.Append(compatibility1);
            settings1.Append(rsids1);
            settings1.Append(mathProperties1);
            settings1.Append(uICompatibleWith97To20031);
            settings1.Append(themeFontLanguages1);
            settings1.Append(colorSchemeMapping1);
            settings1.Append(doNotIncludeSubdocsInStats1);
            settings1.Append(doNotAutoCompressPictures1);
            settings1.Append(decimalSymbol1);
            settings1.Append(listSeparator1);

            documentSettingsPart1.Settings = settings1;
        }

        // Generates content of fontTablePart1.
        private void GenerateFontTablePart1Content(FontTablePart fontTablePart1)
        {
            Fonts fonts1 = new Fonts();
            fonts1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            fonts1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            Font font1 = new Font(){ Name = "Times New Roman" };
            Panose1Number panose1Number1 = new Panose1Number(){ Val = "02020603050405020304" };
            FontCharSet fontCharSet1 = new FontCharSet(){ Val = "CC" };
            FontFamily fontFamily1 = new FontFamily(){ Val = FontFamilyValues.Roman };
            Pitch pitch1 = new Pitch(){ Val = FontPitchValues.Variable };
            FontSignature fontSignature1 = new FontSignature(){ UnicodeSignature0 = "E0002AFF", UnicodeSignature1 = "C0007841", UnicodeSignature2 = "00000009", UnicodeSignature3 = "00000000", CodePageSignature0 = "000001FF", CodePageSignature1 = "00000000" };

            font1.Append(panose1Number1);
            font1.Append(fontCharSet1);
            font1.Append(fontFamily1);
            font1.Append(pitch1);
            font1.Append(fontSignature1);

            Font font2 = new Font(){ Name = "Calibri" };
            Panose1Number panose1Number2 = new Panose1Number(){ Val = "020F0502020204030204" };
            FontCharSet fontCharSet2 = new FontCharSet(){ Val = "CC" };
            FontFamily fontFamily2 = new FontFamily(){ Val = FontFamilyValues.Swiss };
            Pitch pitch2 = new Pitch(){ Val = FontPitchValues.Variable };
            FontSignature fontSignature2 = new FontSignature(){ UnicodeSignature0 = "E10002FF", UnicodeSignature1 = "4000ACFF", UnicodeSignature2 = "00000009", UnicodeSignature3 = "00000000", CodePageSignature0 = "0000019F", CodePageSignature1 = "00000000" };

            font2.Append(panose1Number2);
            font2.Append(fontCharSet2);
            font2.Append(fontFamily2);
            font2.Append(pitch2);
            font2.Append(fontSignature2);

            Font font3 = new Font(){ Name = "Courier New" };
            Panose1Number panose1Number3 = new Panose1Number(){ Val = "02070309020205020404" };
            FontCharSet fontCharSet3 = new FontCharSet(){ Val = "CC" };
            FontFamily fontFamily3 = new FontFamily(){ Val = FontFamilyValues.Modern };
            Pitch pitch3 = new Pitch(){ Val = FontPitchValues.Fixed };
            FontSignature fontSignature3 = new FontSignature(){ UnicodeSignature0 = "E0002AFF", UnicodeSignature1 = "C0007843", UnicodeSignature2 = "00000009", UnicodeSignature3 = "00000000", CodePageSignature0 = "000001FF", CodePageSignature1 = "00000000" };

            font3.Append(panose1Number3);
            font3.Append(fontCharSet3);
            font3.Append(fontFamily3);
            font3.Append(pitch3);
            font3.Append(fontSignature3);

            Font font4 = new Font(){ Name = "Cambria" };
            Panose1Number panose1Number4 = new Panose1Number(){ Val = "02040503050406030204" };
            FontCharSet fontCharSet4 = new FontCharSet(){ Val = "CC" };
            FontFamily fontFamily4 = new FontFamily(){ Val = FontFamilyValues.Roman };
            Pitch pitch4 = new Pitch(){ Val = FontPitchValues.Variable };
            FontSignature fontSignature4 = new FontSignature(){ UnicodeSignature0 = "E00002FF", UnicodeSignature1 = "400004FF", UnicodeSignature2 = "00000000", UnicodeSignature3 = "00000000", CodePageSignature0 = "0000019F", CodePageSignature1 = "00000000" };

            font4.Append(panose1Number4);
            font4.Append(fontCharSet4);
            font4.Append(fontFamily4);
            font4.Append(pitch4);
            font4.Append(fontSignature4);

            fonts1.Append(font1);
            fonts1.Append(font2);
            fonts1.Append(font3);
            fonts1.Append(font4);

            fontTablePart1.Fonts = fonts1;
        }

        // Generates content of styleDefinitionsPart1.
        private void GenerateStyleDefinitionsPart1Content(StyleDefinitionsPart styleDefinitionsPart1)
        {
            Styles styles1 = new Styles();
            styles1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            styles1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            DocDefaults docDefaults1 = new DocDefaults();

            RunPropertiesDefault runPropertiesDefault1 = new RunPropertiesDefault();

            RunPropertiesBaseStyle runPropertiesBaseStyle1 = new RunPropertiesBaseStyle();
            RunFonts runFonts1 = new RunFonts(){ Ascii = "Calibri", HighAnsi = "Calibri", EastAsia = "Calibri", ComplexScript = "Times New Roman" };
            FontSize fontSize1 = new FontSize(){ Val = "22" };
            FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript(){ Val = "22" };
            Languages languages2 = new Languages(){ Val = "ru-RU", EastAsia = "ru-RU", Bidi = "ar-SA" };

            runPropertiesBaseStyle1.Append(runFonts1);
            runPropertiesBaseStyle1.Append(fontSize1);
            runPropertiesBaseStyle1.Append(fontSizeComplexScript2);
            runPropertiesBaseStyle1.Append(languages2);

            runPropertiesDefault1.Append(runPropertiesBaseStyle1);
            ParagraphPropertiesDefault paragraphPropertiesDefault1 = new ParagraphPropertiesDefault();

            docDefaults1.Append(runPropertiesDefault1);
            docDefaults1.Append(paragraphPropertiesDefault1);

            LatentStyles latentStyles1 = new LatentStyles(){ DefaultLockedState = false, DefaultUiPriority = 99, DefaultSemiHidden = true, DefaultUnhideWhenUsed = true, DefaultPrimaryStyle = false, Count = 267 };
            LatentStyleExceptionInfo latentStyleExceptionInfo1 = new LatentStyleExceptionInfo(){ Name = "Normal", Locked = true, UiPriority = 0, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo2 = new LatentStyleExceptionInfo(){ Name = "heading 1", Locked = true, UiPriority = 0, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo3 = new LatentStyleExceptionInfo(){ Name = "heading 2", Locked = true, UiPriority = 0, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo4 = new LatentStyleExceptionInfo(){ Name = "heading 3", Locked = true, UiPriority = 0, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo5 = new LatentStyleExceptionInfo(){ Name = "heading 4", Locked = true, UiPriority = 0, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo6 = new LatentStyleExceptionInfo(){ Name = "heading 5", Locked = true, UiPriority = 0, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo7 = new LatentStyleExceptionInfo(){ Name = "heading 6", Locked = true, UiPriority = 0, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo8 = new LatentStyleExceptionInfo(){ Name = "heading 7", Locked = true, UiPriority = 0, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo9 = new LatentStyleExceptionInfo(){ Name = "heading 8", Locked = true, UiPriority = 0, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo10 = new LatentStyleExceptionInfo(){ Name = "heading 9", Locked = true, UiPriority = 0, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo11 = new LatentStyleExceptionInfo(){ Name = "toc 1", Locked = true, UiPriority = 0, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo12 = new LatentStyleExceptionInfo(){ Name = "toc 2", Locked = true, UiPriority = 0, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo13 = new LatentStyleExceptionInfo(){ Name = "toc 3", Locked = true, UiPriority = 0, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo14 = new LatentStyleExceptionInfo(){ Name = "toc 4", Locked = true, UiPriority = 0, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo15 = new LatentStyleExceptionInfo(){ Name = "toc 5", Locked = true, UiPriority = 0, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo16 = new LatentStyleExceptionInfo(){ Name = "toc 6", Locked = true, UiPriority = 0, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo17 = new LatentStyleExceptionInfo(){ Name = "toc 7", Locked = true, UiPriority = 0, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo18 = new LatentStyleExceptionInfo(){ Name = "toc 8", Locked = true, UiPriority = 0, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo19 = new LatentStyleExceptionInfo(){ Name = "toc 9", Locked = true, UiPriority = 0, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo20 = new LatentStyleExceptionInfo(){ Name = "caption", Locked = true, UiPriority = 0, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo21 = new LatentStyleExceptionInfo(){ Name = "Title", Locked = true, UiPriority = 0, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo22 = new LatentStyleExceptionInfo(){ Name = "Default Paragraph Font", Locked = true, UiPriority = 0, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo23 = new LatentStyleExceptionInfo(){ Name = "Subtitle", Locked = true, UiPriority = 0, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo24 = new LatentStyleExceptionInfo(){ Name = "Strong", Locked = true, UiPriority = 0, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo25 = new LatentStyleExceptionInfo(){ Name = "Emphasis", Locked = true, UiPriority = 0, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo26 = new LatentStyleExceptionInfo(){ Name = "Table Grid", Locked = true, UiPriority = 0, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo27 = new LatentStyleExceptionInfo(){ Name = "Placeholder Text", UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo28 = new LatentStyleExceptionInfo(){ Name = "No Spacing", UiPriority = 1, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo29 = new LatentStyleExceptionInfo(){ Name = "Light Shading", UiPriority = 60, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo30 = new LatentStyleExceptionInfo(){ Name = "Light List", UiPriority = 61, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo31 = new LatentStyleExceptionInfo(){ Name = "Light Grid", UiPriority = 62, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo32 = new LatentStyleExceptionInfo(){ Name = "Medium Shading 1", UiPriority = 63, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo33 = new LatentStyleExceptionInfo(){ Name = "Medium Shading 2", UiPriority = 64, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo34 = new LatentStyleExceptionInfo(){ Name = "Medium List 1", UiPriority = 65, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo35 = new LatentStyleExceptionInfo(){ Name = "Medium List 2", UiPriority = 66, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo36 = new LatentStyleExceptionInfo(){ Name = "Medium Grid 1", UiPriority = 67, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo37 = new LatentStyleExceptionInfo(){ Name = "Medium Grid 2", UiPriority = 68, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo38 = new LatentStyleExceptionInfo(){ Name = "Medium Grid 3", UiPriority = 69, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo39 = new LatentStyleExceptionInfo(){ Name = "Dark List", UiPriority = 70, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo40 = new LatentStyleExceptionInfo(){ Name = "Colorful Shading", UiPriority = 71, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo41 = new LatentStyleExceptionInfo(){ Name = "Colorful List", UiPriority = 72, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo42 = new LatentStyleExceptionInfo(){ Name = "Colorful Grid", UiPriority = 73, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo43 = new LatentStyleExceptionInfo(){ Name = "Light Shading Accent 1", UiPriority = 60, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo44 = new LatentStyleExceptionInfo(){ Name = "Light List Accent 1", UiPriority = 61, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo45 = new LatentStyleExceptionInfo(){ Name = "Light Grid Accent 1", UiPriority = 62, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo46 = new LatentStyleExceptionInfo(){ Name = "Medium Shading 1 Accent 1", UiPriority = 63, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo47 = new LatentStyleExceptionInfo(){ Name = "Medium Shading 2 Accent 1", UiPriority = 64, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo48 = new LatentStyleExceptionInfo(){ Name = "Medium List 1 Accent 1", UiPriority = 65, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo49 = new LatentStyleExceptionInfo(){ Name = "Revision", UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo50 = new LatentStyleExceptionInfo(){ Name = "List Paragraph", UiPriority = 34, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo51 = new LatentStyleExceptionInfo(){ Name = "Quote", UiPriority = 29, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo52 = new LatentStyleExceptionInfo(){ Name = "Intense Quote", UiPriority = 30, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo53 = new LatentStyleExceptionInfo(){ Name = "Medium List 2 Accent 1", UiPriority = 66, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo54 = new LatentStyleExceptionInfo(){ Name = "Medium Grid 1 Accent 1", UiPriority = 67, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo55 = new LatentStyleExceptionInfo(){ Name = "Medium Grid 2 Accent 1", UiPriority = 68, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo56 = new LatentStyleExceptionInfo(){ Name = "Medium Grid 3 Accent 1", UiPriority = 69, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo57 = new LatentStyleExceptionInfo(){ Name = "Dark List Accent 1", UiPriority = 70, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo58 = new LatentStyleExceptionInfo(){ Name = "Colorful Shading Accent 1", UiPriority = 71, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo59 = new LatentStyleExceptionInfo(){ Name = "Colorful List Accent 1", UiPriority = 72, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo60 = new LatentStyleExceptionInfo(){ Name = "Colorful Grid Accent 1", UiPriority = 73, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo61 = new LatentStyleExceptionInfo(){ Name = "Light Shading Accent 2", UiPriority = 60, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo62 = new LatentStyleExceptionInfo(){ Name = "Light List Accent 2", UiPriority = 61, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo63 = new LatentStyleExceptionInfo(){ Name = "Light Grid Accent 2", UiPriority = 62, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo64 = new LatentStyleExceptionInfo(){ Name = "Medium Shading 1 Accent 2", UiPriority = 63, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo65 = new LatentStyleExceptionInfo(){ Name = "Medium Shading 2 Accent 2", UiPriority = 64, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo66 = new LatentStyleExceptionInfo(){ Name = "Medium List 1 Accent 2", UiPriority = 65, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo67 = new LatentStyleExceptionInfo(){ Name = "Medium List 2 Accent 2", UiPriority = 66, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo68 = new LatentStyleExceptionInfo(){ Name = "Medium Grid 1 Accent 2", UiPriority = 67, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo69 = new LatentStyleExceptionInfo(){ Name = "Medium Grid 2 Accent 2", UiPriority = 68, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo70 = new LatentStyleExceptionInfo(){ Name = "Medium Grid 3 Accent 2", UiPriority = 69, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo71 = new LatentStyleExceptionInfo(){ Name = "Dark List Accent 2", UiPriority = 70, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo72 = new LatentStyleExceptionInfo(){ Name = "Colorful Shading Accent 2", UiPriority = 71, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo73 = new LatentStyleExceptionInfo(){ Name = "Colorful List Accent 2", UiPriority = 72, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo74 = new LatentStyleExceptionInfo(){ Name = "Colorful Grid Accent 2", UiPriority = 73, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo75 = new LatentStyleExceptionInfo(){ Name = "Light Shading Accent 3", UiPriority = 60, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo76 = new LatentStyleExceptionInfo(){ Name = "Light List Accent 3", UiPriority = 61, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo77 = new LatentStyleExceptionInfo(){ Name = "Light Grid Accent 3", UiPriority = 62, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo78 = new LatentStyleExceptionInfo(){ Name = "Medium Shading 1 Accent 3", UiPriority = 63, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo79 = new LatentStyleExceptionInfo(){ Name = "Medium Shading 2 Accent 3", UiPriority = 64, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo80 = new LatentStyleExceptionInfo(){ Name = "Medium List 1 Accent 3", UiPriority = 65, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo81 = new LatentStyleExceptionInfo(){ Name = "Medium List 2 Accent 3", UiPriority = 66, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo82 = new LatentStyleExceptionInfo(){ Name = "Medium Grid 1 Accent 3", UiPriority = 67, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo83 = new LatentStyleExceptionInfo(){ Name = "Medium Grid 2 Accent 3", UiPriority = 68, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo84 = new LatentStyleExceptionInfo(){ Name = "Medium Grid 3 Accent 3", UiPriority = 69, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo85 = new LatentStyleExceptionInfo(){ Name = "Dark List Accent 3", UiPriority = 70, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo86 = new LatentStyleExceptionInfo(){ Name = "Colorful Shading Accent 3", UiPriority = 71, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo87 = new LatentStyleExceptionInfo(){ Name = "Colorful List Accent 3", UiPriority = 72, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo88 = new LatentStyleExceptionInfo(){ Name = "Colorful Grid Accent 3", UiPriority = 73, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo89 = new LatentStyleExceptionInfo(){ Name = "Light Shading Accent 4", UiPriority = 60, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo90 = new LatentStyleExceptionInfo(){ Name = "Light List Accent 4", UiPriority = 61, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo91 = new LatentStyleExceptionInfo(){ Name = "Light Grid Accent 4", UiPriority = 62, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo92 = new LatentStyleExceptionInfo(){ Name = "Medium Shading 1 Accent 4", UiPriority = 63, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo93 = new LatentStyleExceptionInfo(){ Name = "Medium Shading 2 Accent 4", UiPriority = 64, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo94 = new LatentStyleExceptionInfo(){ Name = "Medium List 1 Accent 4", UiPriority = 65, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo95 = new LatentStyleExceptionInfo(){ Name = "Medium List 2 Accent 4", UiPriority = 66, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo96 = new LatentStyleExceptionInfo(){ Name = "Medium Grid 1 Accent 4", UiPriority = 67, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo97 = new LatentStyleExceptionInfo(){ Name = "Medium Grid 2 Accent 4", UiPriority = 68, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo98 = new LatentStyleExceptionInfo(){ Name = "Medium Grid 3 Accent 4", UiPriority = 69, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo99 = new LatentStyleExceptionInfo(){ Name = "Dark List Accent 4", UiPriority = 70, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo100 = new LatentStyleExceptionInfo(){ Name = "Colorful Shading Accent 4", UiPriority = 71, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo101 = new LatentStyleExceptionInfo(){ Name = "Colorful List Accent 4", UiPriority = 72, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo102 = new LatentStyleExceptionInfo(){ Name = "Colorful Grid Accent 4", UiPriority = 73, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo103 = new LatentStyleExceptionInfo(){ Name = "Light Shading Accent 5", UiPriority = 60, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo104 = new LatentStyleExceptionInfo(){ Name = "Light List Accent 5", UiPriority = 61, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo105 = new LatentStyleExceptionInfo(){ Name = "Light Grid Accent 5", UiPriority = 62, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo106 = new LatentStyleExceptionInfo(){ Name = "Medium Shading 1 Accent 5", UiPriority = 63, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo107 = new LatentStyleExceptionInfo(){ Name = "Medium Shading 2 Accent 5", UiPriority = 64, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo108 = new LatentStyleExceptionInfo(){ Name = "Medium List 1 Accent 5", UiPriority = 65, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo109 = new LatentStyleExceptionInfo(){ Name = "Medium List 2 Accent 5", UiPriority = 66, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo110 = new LatentStyleExceptionInfo(){ Name = "Medium Grid 1 Accent 5", UiPriority = 67, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo111 = new LatentStyleExceptionInfo(){ Name = "Medium Grid 2 Accent 5", UiPriority = 68, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo112 = new LatentStyleExceptionInfo(){ Name = "Medium Grid 3 Accent 5", UiPriority = 69, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo113 = new LatentStyleExceptionInfo(){ Name = "Dark List Accent 5", UiPriority = 70, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo114 = new LatentStyleExceptionInfo(){ Name = "Colorful Shading Accent 5", UiPriority = 71, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo115 = new LatentStyleExceptionInfo(){ Name = "Colorful List Accent 5", UiPriority = 72, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo116 = new LatentStyleExceptionInfo(){ Name = "Colorful Grid Accent 5", UiPriority = 73, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo117 = new LatentStyleExceptionInfo(){ Name = "Light Shading Accent 6", UiPriority = 60, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo118 = new LatentStyleExceptionInfo(){ Name = "Light List Accent 6", UiPriority = 61, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo119 = new LatentStyleExceptionInfo(){ Name = "Light Grid Accent 6", UiPriority = 62, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo120 = new LatentStyleExceptionInfo(){ Name = "Medium Shading 1 Accent 6", UiPriority = 63, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo121 = new LatentStyleExceptionInfo(){ Name = "Medium Shading 2 Accent 6", UiPriority = 64, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo122 = new LatentStyleExceptionInfo(){ Name = "Medium List 1 Accent 6", UiPriority = 65, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo123 = new LatentStyleExceptionInfo(){ Name = "Medium List 2 Accent 6", UiPriority = 66, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo124 = new LatentStyleExceptionInfo(){ Name = "Medium Grid 1 Accent 6", UiPriority = 67, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo125 = new LatentStyleExceptionInfo(){ Name = "Medium Grid 2 Accent 6", UiPriority = 68, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo126 = new LatentStyleExceptionInfo(){ Name = "Medium Grid 3 Accent 6", UiPriority = 69, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo127 = new LatentStyleExceptionInfo(){ Name = "Dark List Accent 6", UiPriority = 70, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo128 = new LatentStyleExceptionInfo(){ Name = "Colorful Shading Accent 6", UiPriority = 71, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo129 = new LatentStyleExceptionInfo(){ Name = "Colorful List Accent 6", UiPriority = 72, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo130 = new LatentStyleExceptionInfo(){ Name = "Colorful Grid Accent 6", UiPriority = 73, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo131 = new LatentStyleExceptionInfo(){ Name = "Subtle Emphasis", UiPriority = 19, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo132 = new LatentStyleExceptionInfo(){ Name = "Intense Emphasis", UiPriority = 21, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo133 = new LatentStyleExceptionInfo(){ Name = "Subtle Reference", UiPriority = 31, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo134 = new LatentStyleExceptionInfo(){ Name = "Intense Reference", UiPriority = 32, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo135 = new LatentStyleExceptionInfo(){ Name = "Book Title", UiPriority = 33, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo136 = new LatentStyleExceptionInfo(){ Name = "Bibliography", UiPriority = 37 };
            LatentStyleExceptionInfo latentStyleExceptionInfo137 = new LatentStyleExceptionInfo(){ Name = "TOC Heading", UiPriority = 39, PrimaryStyle = true };

            latentStyles1.Append(latentStyleExceptionInfo1);
            latentStyles1.Append(latentStyleExceptionInfo2);
            latentStyles1.Append(latentStyleExceptionInfo3);
            latentStyles1.Append(latentStyleExceptionInfo4);
            latentStyles1.Append(latentStyleExceptionInfo5);
            latentStyles1.Append(latentStyleExceptionInfo6);
            latentStyles1.Append(latentStyleExceptionInfo7);
            latentStyles1.Append(latentStyleExceptionInfo8);
            latentStyles1.Append(latentStyleExceptionInfo9);
            latentStyles1.Append(latentStyleExceptionInfo10);
            latentStyles1.Append(latentStyleExceptionInfo11);
            latentStyles1.Append(latentStyleExceptionInfo12);
            latentStyles1.Append(latentStyleExceptionInfo13);
            latentStyles1.Append(latentStyleExceptionInfo14);
            latentStyles1.Append(latentStyleExceptionInfo15);
            latentStyles1.Append(latentStyleExceptionInfo16);
            latentStyles1.Append(latentStyleExceptionInfo17);
            latentStyles1.Append(latentStyleExceptionInfo18);
            latentStyles1.Append(latentStyleExceptionInfo19);
            latentStyles1.Append(latentStyleExceptionInfo20);
            latentStyles1.Append(latentStyleExceptionInfo21);
            latentStyles1.Append(latentStyleExceptionInfo22);
            latentStyles1.Append(latentStyleExceptionInfo23);
            latentStyles1.Append(latentStyleExceptionInfo24);
            latentStyles1.Append(latentStyleExceptionInfo25);
            latentStyles1.Append(latentStyleExceptionInfo26);
            latentStyles1.Append(latentStyleExceptionInfo27);
            latentStyles1.Append(latentStyleExceptionInfo28);
            latentStyles1.Append(latentStyleExceptionInfo29);
            latentStyles1.Append(latentStyleExceptionInfo30);
            latentStyles1.Append(latentStyleExceptionInfo31);
            latentStyles1.Append(latentStyleExceptionInfo32);
            latentStyles1.Append(latentStyleExceptionInfo33);
            latentStyles1.Append(latentStyleExceptionInfo34);
            latentStyles1.Append(latentStyleExceptionInfo35);
            latentStyles1.Append(latentStyleExceptionInfo36);
            latentStyles1.Append(latentStyleExceptionInfo37);
            latentStyles1.Append(latentStyleExceptionInfo38);
            latentStyles1.Append(latentStyleExceptionInfo39);
            latentStyles1.Append(latentStyleExceptionInfo40);
            latentStyles1.Append(latentStyleExceptionInfo41);
            latentStyles1.Append(latentStyleExceptionInfo42);
            latentStyles1.Append(latentStyleExceptionInfo43);
            latentStyles1.Append(latentStyleExceptionInfo44);
            latentStyles1.Append(latentStyleExceptionInfo45);
            latentStyles1.Append(latentStyleExceptionInfo46);
            latentStyles1.Append(latentStyleExceptionInfo47);
            latentStyles1.Append(latentStyleExceptionInfo48);
            latentStyles1.Append(latentStyleExceptionInfo49);
            latentStyles1.Append(latentStyleExceptionInfo50);
            latentStyles1.Append(latentStyleExceptionInfo51);
            latentStyles1.Append(latentStyleExceptionInfo52);
            latentStyles1.Append(latentStyleExceptionInfo53);
            latentStyles1.Append(latentStyleExceptionInfo54);
            latentStyles1.Append(latentStyleExceptionInfo55);
            latentStyles1.Append(latentStyleExceptionInfo56);
            latentStyles1.Append(latentStyleExceptionInfo57);
            latentStyles1.Append(latentStyleExceptionInfo58);
            latentStyles1.Append(latentStyleExceptionInfo59);
            latentStyles1.Append(latentStyleExceptionInfo60);
            latentStyles1.Append(latentStyleExceptionInfo61);
            latentStyles1.Append(latentStyleExceptionInfo62);
            latentStyles1.Append(latentStyleExceptionInfo63);
            latentStyles1.Append(latentStyleExceptionInfo64);
            latentStyles1.Append(latentStyleExceptionInfo65);
            latentStyles1.Append(latentStyleExceptionInfo66);
            latentStyles1.Append(latentStyleExceptionInfo67);
            latentStyles1.Append(latentStyleExceptionInfo68);
            latentStyles1.Append(latentStyleExceptionInfo69);
            latentStyles1.Append(latentStyleExceptionInfo70);
            latentStyles1.Append(latentStyleExceptionInfo71);
            latentStyles1.Append(latentStyleExceptionInfo72);
            latentStyles1.Append(latentStyleExceptionInfo73);
            latentStyles1.Append(latentStyleExceptionInfo74);
            latentStyles1.Append(latentStyleExceptionInfo75);
            latentStyles1.Append(latentStyleExceptionInfo76);
            latentStyles1.Append(latentStyleExceptionInfo77);
            latentStyles1.Append(latentStyleExceptionInfo78);
            latentStyles1.Append(latentStyleExceptionInfo79);
            latentStyles1.Append(latentStyleExceptionInfo80);
            latentStyles1.Append(latentStyleExceptionInfo81);
            latentStyles1.Append(latentStyleExceptionInfo82);
            latentStyles1.Append(latentStyleExceptionInfo83);
            latentStyles1.Append(latentStyleExceptionInfo84);
            latentStyles1.Append(latentStyleExceptionInfo85);
            latentStyles1.Append(latentStyleExceptionInfo86);
            latentStyles1.Append(latentStyleExceptionInfo87);
            latentStyles1.Append(latentStyleExceptionInfo88);
            latentStyles1.Append(latentStyleExceptionInfo89);
            latentStyles1.Append(latentStyleExceptionInfo90);
            latentStyles1.Append(latentStyleExceptionInfo91);
            latentStyles1.Append(latentStyleExceptionInfo92);
            latentStyles1.Append(latentStyleExceptionInfo93);
            latentStyles1.Append(latentStyleExceptionInfo94);
            latentStyles1.Append(latentStyleExceptionInfo95);
            latentStyles1.Append(latentStyleExceptionInfo96);
            latentStyles1.Append(latentStyleExceptionInfo97);
            latentStyles1.Append(latentStyleExceptionInfo98);
            latentStyles1.Append(latentStyleExceptionInfo99);
            latentStyles1.Append(latentStyleExceptionInfo100);
            latentStyles1.Append(latentStyleExceptionInfo101);
            latentStyles1.Append(latentStyleExceptionInfo102);
            latentStyles1.Append(latentStyleExceptionInfo103);
            latentStyles1.Append(latentStyleExceptionInfo104);
            latentStyles1.Append(latentStyleExceptionInfo105);
            latentStyles1.Append(latentStyleExceptionInfo106);
            latentStyles1.Append(latentStyleExceptionInfo107);
            latentStyles1.Append(latentStyleExceptionInfo108);
            latentStyles1.Append(latentStyleExceptionInfo109);
            latentStyles1.Append(latentStyleExceptionInfo110);
            latentStyles1.Append(latentStyleExceptionInfo111);
            latentStyles1.Append(latentStyleExceptionInfo112);
            latentStyles1.Append(latentStyleExceptionInfo113);
            latentStyles1.Append(latentStyleExceptionInfo114);
            latentStyles1.Append(latentStyleExceptionInfo115);
            latentStyles1.Append(latentStyleExceptionInfo116);
            latentStyles1.Append(latentStyleExceptionInfo117);
            latentStyles1.Append(latentStyleExceptionInfo118);
            latentStyles1.Append(latentStyleExceptionInfo119);
            latentStyles1.Append(latentStyleExceptionInfo120);
            latentStyles1.Append(latentStyleExceptionInfo121);
            latentStyles1.Append(latentStyleExceptionInfo122);
            latentStyles1.Append(latentStyleExceptionInfo123);
            latentStyles1.Append(latentStyleExceptionInfo124);
            latentStyles1.Append(latentStyleExceptionInfo125);
            latentStyles1.Append(latentStyleExceptionInfo126);
            latentStyles1.Append(latentStyleExceptionInfo127);
            latentStyles1.Append(latentStyleExceptionInfo128);
            latentStyles1.Append(latentStyleExceptionInfo129);
            latentStyles1.Append(latentStyleExceptionInfo130);
            latentStyles1.Append(latentStyleExceptionInfo131);
            latentStyles1.Append(latentStyleExceptionInfo132);
            latentStyles1.Append(latentStyleExceptionInfo133);
            latentStyles1.Append(latentStyleExceptionInfo134);
            latentStyles1.Append(latentStyleExceptionInfo135);
            latentStyles1.Append(latentStyleExceptionInfo136);
            latentStyles1.Append(latentStyleExceptionInfo137);

            Style style1 = new Style(){ Type = StyleValues.Paragraph, StyleId = "Normal", Default = true };
            StyleName styleName1 = new StyleName(){ Val = "Normal" };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();
            Rsid rsid46 = new Rsid(){ Val = "00CD6E14" };

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines2 = new SpacingBetweenLines(){ After = "40" };
            Justification justification2 = new Justification(){ Val = JustificationValues.Both };

            styleParagraphProperties1.Append(spacingBetweenLines2);
            styleParagraphProperties1.Append(justification2);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            RunFonts runFonts2 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            FontSize fontSize2 = new FontSize(){ Val = "28" };
            Languages languages3 = new Languages(){ EastAsia = "en-US" };

            styleRunProperties1.Append(runFonts2);
            styleRunProperties1.Append(fontSize2);
            styleRunProperties1.Append(languages3);

            style1.Append(styleName1);
            style1.Append(primaryStyle1);
            style1.Append(rsid46);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleRunProperties1);

            Style style2 = new Style(){ Type = StyleValues.Character, StyleId = "DefaultParagraphFont", Default = true };
            StyleName styleName2 = new StyleName(){ Val = "Default Paragraph Font" };
            UIPriority uIPriority1 = new UIPriority(){ Val = 99 };
            SemiHidden semiHidden1 = new SemiHidden();

            style2.Append(styleName2);
            style2.Append(uIPriority1);
            style2.Append(semiHidden1);

            Style style3 = new Style(){ Type = StyleValues.Table, StyleId = "TableNormal", Default = true };
            StyleName styleName3 = new StyleName(){ Val = "Normal Table" };
            UIPriority uIPriority2 = new UIPriority(){ Val = 99 };
            SemiHidden semiHidden2 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed1 = new UnhideWhenUsed();
            PrimaryStyle primaryStyle2 = new PrimaryStyle();

            StyleTableProperties styleTableProperties1 = new StyleTableProperties();
            TableIndentation tableIndentation1 = new TableIndentation(){ Width = 0, Type = TableWidthUnitValues.Dxa };

            TableCellMarginDefault tableCellMarginDefault1 = new TableCellMarginDefault();
            TopMargin topMargin1 = new TopMargin(){ Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellLeftMargin tableCellLeftMargin1 = new TableCellLeftMargin(){ Width = 108, Type = TableWidthValues.Dxa };
            BottomMargin bottomMargin1 = new BottomMargin(){ Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellRightMargin tableCellRightMargin1 = new TableCellRightMargin(){ Width = 108, Type = TableWidthValues.Dxa };

            tableCellMarginDefault1.Append(topMargin1);
            tableCellMarginDefault1.Append(tableCellLeftMargin1);
            tableCellMarginDefault1.Append(bottomMargin1);
            tableCellMarginDefault1.Append(tableCellRightMargin1);

            styleTableProperties1.Append(tableIndentation1);
            styleTableProperties1.Append(tableCellMarginDefault1);

            style3.Append(styleName3);
            style3.Append(uIPriority2);
            style3.Append(semiHidden2);
            style3.Append(unhideWhenUsed1);
            style3.Append(primaryStyle2);
            style3.Append(styleTableProperties1);

            Style style4 = new Style(){ Type = StyleValues.Numbering, StyleId = "NoList", Default = true };
            StyleName styleName4 = new StyleName(){ Val = "No List" };
            UIPriority uIPriority3 = new UIPriority(){ Val = 99 };
            SemiHidden semiHidden3 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed2 = new UnhideWhenUsed();

            style4.Append(styleName4);
            style4.Append(uIPriority3);
            style4.Append(semiHidden3);
            style4.Append(unhideWhenUsed2);

            Style style5 = new Style(){ Type = StyleValues.Character, StyleId = "pl-k", CustomStyle = true };
            StyleName styleName5 = new StyleName(){ Val = "pl-k" };
            BasedOn basedOn1 = new BasedOn(){ Val = "DefaultParagraphFont" };
            UIPriority uIPriority4 = new UIPriority(){ Val = 99 };
            Rsid rsid47 = new Rsid(){ Val = "007D2F02" };

            StyleRunProperties styleRunProperties2 = new StyleRunProperties();
            RunFonts runFonts3 = new RunFonts(){ ComplexScript = "Times New Roman" };

            styleRunProperties2.Append(runFonts3);

            style5.Append(styleName5);
            style5.Append(basedOn1);
            style5.Append(uIPriority4);
            style5.Append(rsid47);
            style5.Append(styleRunProperties2);

            Style style6 = new Style(){ Type = StyleValues.Character, StyleId = "pl-en", CustomStyle = true };
            StyleName styleName6 = new StyleName(){ Val = "pl-en" };
            BasedOn basedOn2 = new BasedOn(){ Val = "DefaultParagraphFont" };
            UIPriority uIPriority5 = new UIPriority(){ Val = 99 };
            Rsid rsid48 = new Rsid(){ Val = "007D2F02" };

            StyleRunProperties styleRunProperties3 = new StyleRunProperties();
            RunFonts runFonts4 = new RunFonts(){ ComplexScript = "Times New Roman" };

            styleRunProperties3.Append(runFonts4);

            style6.Append(styleName6);
            style6.Append(basedOn2);
            style6.Append(uIPriority5);
            style6.Append(rsid48);
            style6.Append(styleRunProperties3);

            Style style7 = new Style(){ Type = StyleValues.Character, StyleId = "pl-smi", CustomStyle = true };
            StyleName styleName7 = new StyleName(){ Val = "pl-smi" };
            BasedOn basedOn3 = new BasedOn(){ Val = "DefaultParagraphFont" };
            UIPriority uIPriority6 = new UIPriority(){ Val = 99 };
            Rsid rsid49 = new Rsid(){ Val = "007D2F02" };

            StyleRunProperties styleRunProperties4 = new StyleRunProperties();
            RunFonts runFonts5 = new RunFonts(){ ComplexScript = "Times New Roman" };

            styleRunProperties4.Append(runFonts5);

            style7.Append(styleName7);
            style7.Append(basedOn3);
            style7.Append(uIPriority6);
            style7.Append(rsid49);
            style7.Append(styleRunProperties4);

            Style style8 = new Style(){ Type = StyleValues.Character, StyleId = "pl-c1", CustomStyle = true };
            StyleName styleName8 = new StyleName(){ Val = "pl-c1" };
            BasedOn basedOn4 = new BasedOn(){ Val = "DefaultParagraphFont" };
            UIPriority uIPriority7 = new UIPriority(){ Val = 99 };
            Rsid rsid50 = new Rsid(){ Val = "007D2F02" };

            StyleRunProperties styleRunProperties5 = new StyleRunProperties();
            RunFonts runFonts6 = new RunFonts(){ ComplexScript = "Times New Roman" };

            styleRunProperties5.Append(runFonts6);

            style8.Append(styleName8);
            style8.Append(basedOn4);
            style8.Append(uIPriority7);
            style8.Append(rsid50);
            style8.Append(styleRunProperties5);

            Style style9 = new Style(){ Type = StyleValues.Character, StyleId = "pl-s", CustomStyle = true };
            StyleName styleName9 = new StyleName(){ Val = "pl-s" };
            BasedOn basedOn5 = new BasedOn(){ Val = "DefaultParagraphFont" };
            UIPriority uIPriority8 = new UIPriority(){ Val = 99 };
            Rsid rsid51 = new Rsid(){ Val = "007D2F02" };

            StyleRunProperties styleRunProperties6 = new StyleRunProperties();
            RunFonts runFonts7 = new RunFonts(){ ComplexScript = "Times New Roman" };

            styleRunProperties6.Append(runFonts7);

            style9.Append(styleName9);
            style9.Append(basedOn5);
            style9.Append(uIPriority8);
            style9.Append(rsid51);
            style9.Append(styleRunProperties6);

            Style style10 = new Style(){ Type = StyleValues.Character, StyleId = "pl-pds", CustomStyle = true };
            StyleName styleName10 = new StyleName(){ Val = "pl-pds" };
            BasedOn basedOn6 = new BasedOn(){ Val = "DefaultParagraphFont" };
            UIPriority uIPriority9 = new UIPriority(){ Val = 99 };
            Rsid rsid52 = new Rsid(){ Val = "007D2F02" };

            StyleRunProperties styleRunProperties7 = new StyleRunProperties();
            RunFonts runFonts8 = new RunFonts(){ ComplexScript = "Times New Roman" };

            styleRunProperties7.Append(runFonts8);

            style10.Append(styleName10);
            style10.Append(basedOn6);
            style10.Append(uIPriority9);
            style10.Append(rsid52);
            style10.Append(styleRunProperties7);

            Style style11 = new Style(){ Type = StyleValues.Table, StyleId = "TableGrid" };
            StyleName styleName11 = new StyleName(){ Val = "Table Grid" };
            BasedOn basedOn7 = new BasedOn(){ Val = "TableNormal" };
            UIPriority uIPriority10 = new UIPriority(){ Val = 99 };
            Rsid rsid53 = new Rsid(){ Val = "007D2F02" };

            StyleRunProperties styleRunProperties8 = new StyleRunProperties();
            FontSize fontSize3 = new FontSize(){ Val = "20" };
            FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript(){ Val = "20" };

            styleRunProperties8.Append(fontSize3);
            styleRunProperties8.Append(fontSizeComplexScript3);

            StyleTableProperties styleTableProperties2 = new StyleTableProperties();
            TableIndentation tableIndentation2 = new TableIndentation(){ Width = 0, Type = TableWidthUnitValues.Dxa };

            TableBorders tableBorders1 = new TableBorders();
            TopBorder topBorder1 = new TopBorder(){ Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            LeftBorder leftBorder1 = new LeftBorder(){ Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder1 = new BottomBorder(){ Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder1 = new RightBorder(){ Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder1 = new InsideHorizontalBorder(){ Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideVerticalBorder insideVerticalBorder1 = new InsideVerticalBorder(){ Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableBorders1.Append(topBorder1);
            tableBorders1.Append(leftBorder1);
            tableBorders1.Append(bottomBorder1);
            tableBorders1.Append(rightBorder1);
            tableBorders1.Append(insideHorizontalBorder1);
            tableBorders1.Append(insideVerticalBorder1);

            TableCellMarginDefault tableCellMarginDefault2 = new TableCellMarginDefault();
            TopMargin topMargin2 = new TopMargin(){ Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellLeftMargin tableCellLeftMargin2 = new TableCellLeftMargin(){ Width = 108, Type = TableWidthValues.Dxa };
            BottomMargin bottomMargin2 = new BottomMargin(){ Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellRightMargin tableCellRightMargin2 = new TableCellRightMargin(){ Width = 108, Type = TableWidthValues.Dxa };

            tableCellMarginDefault2.Append(topMargin2);
            tableCellMarginDefault2.Append(tableCellLeftMargin2);
            tableCellMarginDefault2.Append(bottomMargin2);
            tableCellMarginDefault2.Append(tableCellRightMargin2);

            styleTableProperties2.Append(tableIndentation2);
            styleTableProperties2.Append(tableBorders1);
            styleTableProperties2.Append(tableCellMarginDefault2);

            style11.Append(styleName11);
            style11.Append(basedOn7);
            style11.Append(uIPriority10);
            style11.Append(rsid53);
            style11.Append(styleRunProperties8);
            style11.Append(styleTableProperties2);

            Style style12 = new Style(){ Type = StyleValues.Paragraph, StyleId = "Header" };
            StyleName styleName12 = new StyleName(){ Val = "header" };
            BasedOn basedOn8 = new BasedOn(){ Val = "Normal" };
            LinkedStyle linkedStyle1 = new LinkedStyle(){ Val = "HeaderChar" };
            UIPriority uIPriority11 = new UIPriority(){ Val = 99 };
            Rsid rsid54 = new Rsid(){ Val = "007D2F02" };

            StyleParagraphProperties styleParagraphProperties2 = new StyleParagraphProperties();

            Tabs tabs1 = new Tabs();
            TabStop tabStop1 = new TabStop(){ Val = TabStopValues.Center, Position = 4677 };
            TabStop tabStop2 = new TabStop(){ Val = TabStopValues.Right, Position = 9355 };

            tabs1.Append(tabStop1);
            tabs1.Append(tabStop2);
            SpacingBetweenLines spacingBetweenLines3 = new SpacingBetweenLines(){ After = "0" };

            styleParagraphProperties2.Append(tabs1);
            styleParagraphProperties2.Append(spacingBetweenLines3);

            style12.Append(styleName12);
            style12.Append(basedOn8);
            style12.Append(linkedStyle1);
            style12.Append(uIPriority11);
            style12.Append(rsid54);
            style12.Append(styleParagraphProperties2);

            Style style13 = new Style(){ Type = StyleValues.Character, StyleId = "HeaderChar", CustomStyle = true };
            StyleName styleName13 = new StyleName(){ Val = "Header Char" };
            BasedOn basedOn9 = new BasedOn(){ Val = "DefaultParagraphFont" };
            LinkedStyle linkedStyle2 = new LinkedStyle(){ Val = "Header" };
            UIPriority uIPriority12 = new UIPriority(){ Val = 99 };
            Locked locked1 = new Locked();
            Rsid rsid55 = new Rsid(){ Val = "007D2F02" };

            StyleRunProperties styleRunProperties9 = new StyleRunProperties();
            RunFonts runFonts9 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
            FontSize fontSize4 = new FontSize(){ Val = "28" };

            styleRunProperties9.Append(runFonts9);
            styleRunProperties9.Append(fontSize4);

            style13.Append(styleName13);
            style13.Append(basedOn9);
            style13.Append(linkedStyle2);
            style13.Append(uIPriority12);
            style13.Append(locked1);
            style13.Append(rsid55);
            style13.Append(styleRunProperties9);

            Style style14 = new Style(){ Type = StyleValues.Paragraph, StyleId = "Footer" };
            StyleName styleName14 = new StyleName(){ Val = "footer" };
            BasedOn basedOn10 = new BasedOn(){ Val = "Normal" };
            LinkedStyle linkedStyle3 = new LinkedStyle(){ Val = "FooterChar" };
            UIPriority uIPriority13 = new UIPriority(){ Val = 99 };
            Rsid rsid56 = new Rsid(){ Val = "007D2F02" };

            StyleParagraphProperties styleParagraphProperties3 = new StyleParagraphProperties();

            Tabs tabs2 = new Tabs();
            TabStop tabStop3 = new TabStop(){ Val = TabStopValues.Center, Position = 4677 };
            TabStop tabStop4 = new TabStop(){ Val = TabStopValues.Right, Position = 9355 };

            tabs2.Append(tabStop3);
            tabs2.Append(tabStop4);
            SpacingBetweenLines spacingBetweenLines4 = new SpacingBetweenLines(){ After = "0" };

            styleParagraphProperties3.Append(tabs2);
            styleParagraphProperties3.Append(spacingBetweenLines4);

            style14.Append(styleName14);
            style14.Append(basedOn10);
            style14.Append(linkedStyle3);
            style14.Append(uIPriority13);
            style14.Append(rsid56);
            style14.Append(styleParagraphProperties3);

            Style style15 = new Style(){ Type = StyleValues.Character, StyleId = "FooterChar", CustomStyle = true };
            StyleName styleName15 = new StyleName(){ Val = "Footer Char" };
            BasedOn basedOn11 = new BasedOn(){ Val = "DefaultParagraphFont" };
            LinkedStyle linkedStyle4 = new LinkedStyle(){ Val = "Footer" };
            UIPriority uIPriority14 = new UIPriority(){ Val = 99 };
            Locked locked2 = new Locked();
            Rsid rsid57 = new Rsid(){ Val = "007D2F02" };

            StyleRunProperties styleRunProperties10 = new StyleRunProperties();
            RunFonts runFonts10 = new RunFonts(){ Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
            FontSize fontSize5 = new FontSize(){ Val = "28" };

            styleRunProperties10.Append(runFonts10);
            styleRunProperties10.Append(fontSize5);

            style15.Append(styleName15);
            style15.Append(basedOn11);
            style15.Append(linkedStyle4);
            style15.Append(uIPriority14);
            style15.Append(locked2);
            style15.Append(rsid57);
            style15.Append(styleRunProperties10);

            Style style16 = new Style(){ Type = StyleValues.Paragraph, StyleId = "ListParagraph" };
            StyleName styleName16 = new StyleName(){ Val = "List Paragraph" };
            BasedOn basedOn12 = new BasedOn(){ Val = "Normal" };
            UIPriority uIPriority15 = new UIPriority(){ Val = 99 };
            PrimaryStyle primaryStyle3 = new PrimaryStyle();
            Rsid rsid58 = new Rsid(){ Val = "00107029" };

            StyleParagraphProperties styleParagraphProperties4 = new StyleParagraphProperties();
            Indentation indentation1 = new Indentation(){ Left = "720" };
            ContextualSpacing contextualSpacing1 = new ContextualSpacing();

            styleParagraphProperties4.Append(indentation1);
            styleParagraphProperties4.Append(contextualSpacing1);

            style16.Append(styleName16);
            style16.Append(basedOn12);
            style16.Append(uIPriority15);
            style16.Append(primaryStyle3);
            style16.Append(rsid58);
            style16.Append(styleParagraphProperties4);

            Style style17 = new Style(){ Type = StyleValues.Paragraph, StyleId = "NormalWeb" };
            StyleName styleName17 = new StyleName(){ Val = "Normal (Web)" };
            BasedOn basedOn13 = new BasedOn(){ Val = "Normal" };
            UIPriority uIPriority16 = new UIPriority(){ Val = 99 };
            SemiHidden semiHidden4 = new SemiHidden();
            Rsid rsid59 = new Rsid(){ Val = "00071E5A" };

            StyleParagraphProperties styleParagraphProperties5 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines5 = new SpacingBetweenLines(){ Before = "100", BeforeAutoSpacing = true, After = "100", AfterAutoSpacing = true };
            Justification justification3 = new Justification(){ Val = JustificationValues.Left };

            styleParagraphProperties5.Append(spacingBetweenLines5);
            styleParagraphProperties5.Append(justification3);

            StyleRunProperties styleRunProperties11 = new StyleRunProperties();
            RunFonts runFonts11 = new RunFonts(){ EastAsia = "Times New Roman" };
            FontSize fontSize6 = new FontSize(){ Val = "24" };
            FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript(){ Val = "24" };
            Languages languages4 = new Languages(){ EastAsia = "ru-RU" };

            styleRunProperties11.Append(runFonts11);
            styleRunProperties11.Append(fontSize6);
            styleRunProperties11.Append(fontSizeComplexScript4);
            styleRunProperties11.Append(languages4);

            style17.Append(styleName17);
            style17.Append(basedOn13);
            style17.Append(uIPriority16);
            style17.Append(semiHidden4);
            style17.Append(rsid59);
            style17.Append(styleParagraphProperties5);
            style17.Append(styleRunProperties11);

            Style style18 = new Style(){ Type = StyleValues.Paragraph, StyleId = "HTMLPreformatted" };
            StyleName styleName18 = new StyleName(){ Val = "HTML Preformatted" };
            BasedOn basedOn14 = new BasedOn(){ Val = "Normal" };
            LinkedStyle linkedStyle5 = new LinkedStyle(){ Val = "HTMLPreformattedChar" };
            UIPriority uIPriority17 = new UIPriority(){ Val = 99 };
            SemiHidden semiHidden5 = new SemiHidden();
            Rsid rsid60 = new Rsid(){ Val = "00091877" };

            StyleParagraphProperties styleParagraphProperties6 = new StyleParagraphProperties();

            Tabs tabs3 = new Tabs();
            TabStop tabStop5 = new TabStop(){ Val = TabStopValues.Left, Position = 916 };
            TabStop tabStop6 = new TabStop(){ Val = TabStopValues.Left, Position = 1832 };
            TabStop tabStop7 = new TabStop(){ Val = TabStopValues.Left, Position = 2748 };
            TabStop tabStop8 = new TabStop(){ Val = TabStopValues.Left, Position = 3664 };
            TabStop tabStop9 = new TabStop(){ Val = TabStopValues.Left, Position = 4580 };
            TabStop tabStop10 = new TabStop(){ Val = TabStopValues.Left, Position = 5496 };
            TabStop tabStop11 = new TabStop(){ Val = TabStopValues.Left, Position = 6412 };
            TabStop tabStop12 = new TabStop(){ Val = TabStopValues.Left, Position = 7328 };
            TabStop tabStop13 = new TabStop(){ Val = TabStopValues.Left, Position = 8244 };
            TabStop tabStop14 = new TabStop(){ Val = TabStopValues.Left, Position = 9160 };
            TabStop tabStop15 = new TabStop(){ Val = TabStopValues.Left, Position = 10076 };
            TabStop tabStop16 = new TabStop(){ Val = TabStopValues.Left, Position = 10992 };
            TabStop tabStop17 = new TabStop(){ Val = TabStopValues.Left, Position = 11908 };
            TabStop tabStop18 = new TabStop(){ Val = TabStopValues.Left, Position = 12824 };
            TabStop tabStop19 = new TabStop(){ Val = TabStopValues.Left, Position = 13740 };
            TabStop tabStop20 = new TabStop(){ Val = TabStopValues.Left, Position = 14656 };

            tabs3.Append(tabStop5);
            tabs3.Append(tabStop6);
            tabs3.Append(tabStop7);
            tabs3.Append(tabStop8);
            tabs3.Append(tabStop9);
            tabs3.Append(tabStop10);
            tabs3.Append(tabStop11);
            tabs3.Append(tabStop12);
            tabs3.Append(tabStop13);
            tabs3.Append(tabStop14);
            tabs3.Append(tabStop15);
            tabs3.Append(tabStop16);
            tabs3.Append(tabStop17);
            tabs3.Append(tabStop18);
            tabs3.Append(tabStop19);
            tabs3.Append(tabStop20);
            SpacingBetweenLines spacingBetweenLines6 = new SpacingBetweenLines(){ After = "0" };
            Justification justification4 = new Justification(){ Val = JustificationValues.Left };

            styleParagraphProperties6.Append(tabs3);
            styleParagraphProperties6.Append(spacingBetweenLines6);
            styleParagraphProperties6.Append(justification4);

            StyleRunProperties styleRunProperties12 = new StyleRunProperties();
            RunFonts runFonts12 = new RunFonts(){ Ascii = "Courier New", HighAnsi = "Courier New", EastAsia = "Times New Roman", ComplexScript = "Courier New" };
            FontSize fontSize7 = new FontSize(){ Val = "20" };
            FontSizeComplexScript fontSizeComplexScript5 = new FontSizeComplexScript(){ Val = "20" };
            Languages languages5 = new Languages(){ EastAsia = "ru-RU" };

            styleRunProperties12.Append(runFonts12);
            styleRunProperties12.Append(fontSize7);
            styleRunProperties12.Append(fontSizeComplexScript5);
            styleRunProperties12.Append(languages5);

            style18.Append(styleName18);
            style18.Append(basedOn14);
            style18.Append(linkedStyle5);
            style18.Append(uIPriority17);
            style18.Append(semiHidden5);
            style18.Append(rsid60);
            style18.Append(styleParagraphProperties6);
            style18.Append(styleRunProperties12);

            Style style19 = new Style(){ Type = StyleValues.Character, StyleId = "HTMLPreformattedChar", CustomStyle = true };
            StyleName styleName19 = new StyleName(){ Val = "HTML Preformatted Char" };
            BasedOn basedOn15 = new BasedOn(){ Val = "DefaultParagraphFont" };
            LinkedStyle linkedStyle6 = new LinkedStyle(){ Val = "HTMLPreformatted" };
            UIPriority uIPriority18 = new UIPriority(){ Val = 99 };
            SemiHidden semiHidden6 = new SemiHidden();
            Locked locked3 = new Locked();
            Rsid rsid61 = new Rsid(){ Val = "00091877" };

            StyleRunProperties styleRunProperties13 = new StyleRunProperties();
            RunFonts runFonts13 = new RunFonts(){ Ascii = "Courier New", HighAnsi = "Courier New", ComplexScript = "Courier New" };
            FontSize fontSize8 = new FontSize(){ Val = "20" };
            FontSizeComplexScript fontSizeComplexScript6 = new FontSizeComplexScript(){ Val = "20" };
            Languages languages6 = new Languages(){ EastAsia = "ru-RU" };

            styleRunProperties13.Append(runFonts13);
            styleRunProperties13.Append(fontSize8);
            styleRunProperties13.Append(fontSizeComplexScript6);
            styleRunProperties13.Append(languages6);

            style19.Append(styleName19);
            style19.Append(basedOn15);
            style19.Append(linkedStyle6);
            style19.Append(uIPriority18);
            style19.Append(semiHidden6);
            style19.Append(locked3);
            style19.Append(rsid61);
            style19.Append(styleRunProperties13);

            Style style20 = new Style(){ Type = StyleValues.Character, StyleId = "kw4", CustomStyle = true };
            StyleName styleName20 = new StyleName(){ Val = "kw4" };
            BasedOn basedOn16 = new BasedOn(){ Val = "DefaultParagraphFont" };
            UIPriority uIPriority19 = new UIPriority(){ Val = 99 };
            Rsid rsid62 = new Rsid(){ Val = "00091877" };

            StyleRunProperties styleRunProperties14 = new StyleRunProperties();
            RunFonts runFonts14 = new RunFonts(){ ComplexScript = "Times New Roman" };

            styleRunProperties14.Append(runFonts14);

            style20.Append(styleName20);
            style20.Append(basedOn16);
            style20.Append(uIPriority19);
            style20.Append(rsid62);
            style20.Append(styleRunProperties14);

            Style style21 = new Style(){ Type = StyleValues.Character, StyleId = "sy0", CustomStyle = true };
            StyleName styleName21 = new StyleName(){ Val = "sy0" };
            BasedOn basedOn17 = new BasedOn(){ Val = "DefaultParagraphFont" };
            UIPriority uIPriority20 = new UIPriority(){ Val = 99 };
            Rsid rsid63 = new Rsid(){ Val = "00091877" };

            StyleRunProperties styleRunProperties15 = new StyleRunProperties();
            RunFonts runFonts15 = new RunFonts(){ ComplexScript = "Times New Roman" };

            styleRunProperties15.Append(runFonts15);

            style21.Append(styleName21);
            style21.Append(basedOn17);
            style21.Append(uIPriority20);
            style21.Append(rsid63);
            style21.Append(styleRunProperties15);

            Style style22 = new Style(){ Type = StyleValues.Character, StyleId = "kw1", CustomStyle = true };
            StyleName styleName22 = new StyleName(){ Val = "kw1" };
            BasedOn basedOn18 = new BasedOn(){ Val = "DefaultParagraphFont" };
            UIPriority uIPriority21 = new UIPriority(){ Val = 99 };
            Rsid rsid64 = new Rsid(){ Val = "00091877" };

            StyleRunProperties styleRunProperties16 = new StyleRunProperties();
            RunFonts runFonts16 = new RunFonts(){ ComplexScript = "Times New Roman" };

            styleRunProperties16.Append(runFonts16);

            style22.Append(styleName22);
            style22.Append(basedOn18);
            style22.Append(uIPriority21);
            style22.Append(rsid64);
            style22.Append(styleRunProperties16);

            Style style23 = new Style(){ Type = StyleValues.Character, StyleId = "br0", CustomStyle = true };
            StyleName styleName23 = new StyleName(){ Val = "br0" };
            BasedOn basedOn19 = new BasedOn(){ Val = "DefaultParagraphFont" };
            UIPriority uIPriority22 = new UIPriority(){ Val = 99 };
            Rsid rsid65 = new Rsid(){ Val = "00091877" };

            StyleRunProperties styleRunProperties17 = new StyleRunProperties();
            RunFonts runFonts17 = new RunFonts(){ ComplexScript = "Times New Roman" };

            styleRunProperties17.Append(runFonts17);

            style23.Append(styleName23);
            style23.Append(basedOn19);
            style23.Append(uIPriority22);
            style23.Append(rsid65);
            style23.Append(styleRunProperties17);

            Style style24 = new Style(){ Type = StyleValues.Character, StyleId = "nu0", CustomStyle = true };
            StyleName styleName24 = new StyleName(){ Val = "nu0" };
            BasedOn basedOn20 = new BasedOn(){ Val = "DefaultParagraphFont" };
            UIPriority uIPriority23 = new UIPriority(){ Val = 99 };
            Rsid rsid66 = new Rsid(){ Val = "00091877" };

            StyleRunProperties styleRunProperties18 = new StyleRunProperties();
            RunFonts runFonts18 = new RunFonts(){ ComplexScript = "Times New Roman" };

            styleRunProperties18.Append(runFonts18);

            style24.Append(styleName24);
            style24.Append(basedOn20);
            style24.Append(uIPriority23);
            style24.Append(rsid66);
            style24.Append(styleRunProperties18);

            Style style25 = new Style(){ Type = StyleValues.Character, StyleId = "me1", CustomStyle = true };
            StyleName styleName25 = new StyleName(){ Val = "me1" };
            BasedOn basedOn21 = new BasedOn(){ Val = "DefaultParagraphFont" };
            UIPriority uIPriority24 = new UIPriority(){ Val = 99 };
            Rsid rsid67 = new Rsid(){ Val = "00091877" };

            StyleRunProperties styleRunProperties19 = new StyleRunProperties();
            RunFonts runFonts19 = new RunFonts(){ ComplexScript = "Times New Roman" };

            styleRunProperties19.Append(runFonts19);

            style25.Append(styleName25);
            style25.Append(basedOn21);
            style25.Append(uIPriority24);
            style25.Append(rsid67);
            style25.Append(styleRunProperties19);

            Style style26 = new Style(){ Type = StyleValues.Character, StyleId = "mwe-math-mathml-inlinemwe-math-mathml-a11y", CustomStyle = true };
            StyleName styleName26 = new StyleName(){ Val = "mwe-math-mathml-inline mwe-math-mathml-a11y" };
            BasedOn basedOn22 = new BasedOn(){ Val = "DefaultParagraphFont" };
            UIPriority uIPriority25 = new UIPriority(){ Val = 99 };
            Rsid rsid68 = new Rsid(){ Val = "00FA7479" };

            StyleRunProperties styleRunProperties20 = new StyleRunProperties();
            RunFonts runFonts20 = new RunFonts(){ ComplexScript = "Times New Roman" };

            styleRunProperties20.Append(runFonts20);

            style26.Append(styleName26);
            style26.Append(basedOn22);
            style26.Append(uIPriority25);
            style26.Append(rsid68);
            style26.Append(styleRunProperties20);

            Style style27 = new Style(){ Type = StyleValues.Character, StyleId = "crayon-m", CustomStyle = true };
            StyleName styleName27 = new StyleName(){ Val = "crayon-m" };
            BasedOn basedOn23 = new BasedOn(){ Val = "DefaultParagraphFont" };
            UIPriority uIPriority26 = new UIPriority(){ Val = 99 };
            Rsid rsid69 = new Rsid(){ Val = "00FA7479" };

            StyleRunProperties styleRunProperties21 = new StyleRunProperties();
            RunFonts runFonts21 = new RunFonts(){ ComplexScript = "Times New Roman" };

            styleRunProperties21.Append(runFonts21);

            style27.Append(styleName27);
            style27.Append(basedOn23);
            style27.Append(uIPriority26);
            style27.Append(rsid69);
            style27.Append(styleRunProperties21);

            Style style28 = new Style(){ Type = StyleValues.Character, StyleId = "crayon-h", CustomStyle = true };
            StyleName styleName28 = new StyleName(){ Val = "crayon-h" };
            BasedOn basedOn24 = new BasedOn(){ Val = "DefaultParagraphFont" };
            UIPriority uIPriority27 = new UIPriority(){ Val = 99 };
            Rsid rsid70 = new Rsid(){ Val = "00FA7479" };

            StyleRunProperties styleRunProperties22 = new StyleRunProperties();
            RunFonts runFonts22 = new RunFonts(){ ComplexScript = "Times New Roman" };

            styleRunProperties22.Append(runFonts22);

            style28.Append(styleName28);
            style28.Append(basedOn24);
            style28.Append(uIPriority27);
            style28.Append(rsid70);
            style28.Append(styleRunProperties22);

            Style style29 = new Style(){ Type = StyleValues.Character, StyleId = "crayon-t", CustomStyle = true };
            StyleName styleName29 = new StyleName(){ Val = "crayon-t" };
            BasedOn basedOn25 = new BasedOn(){ Val = "DefaultParagraphFont" };
            UIPriority uIPriority28 = new UIPriority(){ Val = 99 };
            Rsid rsid71 = new Rsid(){ Val = "00FA7479" };

            StyleRunProperties styleRunProperties23 = new StyleRunProperties();
            RunFonts runFonts23 = new RunFonts(){ ComplexScript = "Times New Roman" };

            styleRunProperties23.Append(runFonts23);

            style29.Append(styleName29);
            style29.Append(basedOn25);
            style29.Append(uIPriority28);
            style29.Append(rsid71);
            style29.Append(styleRunProperties23);

            Style style30 = new Style(){ Type = StyleValues.Character, StyleId = "crayon-sy", CustomStyle = true };
            StyleName styleName30 = new StyleName(){ Val = "crayon-sy" };
            BasedOn basedOn26 = new BasedOn(){ Val = "DefaultParagraphFont" };
            UIPriority uIPriority29 = new UIPriority(){ Val = 99 };
            Rsid rsid72 = new Rsid(){ Val = "00FA7479" };

            StyleRunProperties styleRunProperties24 = new StyleRunProperties();
            RunFonts runFonts24 = new RunFonts(){ ComplexScript = "Times New Roman" };

            styleRunProperties24.Append(runFonts24);

            style30.Append(styleName30);
            style30.Append(basedOn26);
            style30.Append(uIPriority29);
            style30.Append(rsid72);
            style30.Append(styleRunProperties24);

            Style style31 = new Style(){ Type = StyleValues.Character, StyleId = "crayon-e", CustomStyle = true };
            StyleName styleName31 = new StyleName(){ Val = "crayon-e" };
            BasedOn basedOn27 = new BasedOn(){ Val = "DefaultParagraphFont" };
            UIPriority uIPriority30 = new UIPriority(){ Val = 99 };
            Rsid rsid73 = new Rsid(){ Val = "00FA7479" };

            StyleRunProperties styleRunProperties25 = new StyleRunProperties();
            RunFonts runFonts25 = new RunFonts(){ ComplexScript = "Times New Roman" };

            styleRunProperties25.Append(runFonts25);

            style31.Append(styleName31);
            style31.Append(basedOn27);
            style31.Append(uIPriority30);
            style31.Append(rsid73);
            style31.Append(styleRunProperties25);

            Style style32 = new Style(){ Type = StyleValues.Character, StyleId = "crayon-v", CustomStyle = true };
            StyleName styleName32 = new StyleName(){ Val = "crayon-v" };
            BasedOn basedOn28 = new BasedOn(){ Val = "DefaultParagraphFont" };
            UIPriority uIPriority31 = new UIPriority(){ Val = 99 };
            Rsid rsid74 = new Rsid(){ Val = "00FA7479" };

            StyleRunProperties styleRunProperties26 = new StyleRunProperties();
            RunFonts runFonts26 = new RunFonts(){ ComplexScript = "Times New Roman" };

            styleRunProperties26.Append(runFonts26);

            style32.Append(styleName32);
            style32.Append(basedOn28);
            style32.Append(uIPriority31);
            style32.Append(rsid74);
            style32.Append(styleRunProperties26);

            Style style33 = new Style(){ Type = StyleValues.Character, StyleId = "crayon-st", CustomStyle = true };
            StyleName styleName33 = new StyleName(){ Val = "crayon-st" };
            BasedOn basedOn29 = new BasedOn(){ Val = "DefaultParagraphFont" };
            UIPriority uIPriority32 = new UIPriority(){ Val = 99 };
            Rsid rsid75 = new Rsid(){ Val = "00FA7479" };

            StyleRunProperties styleRunProperties27 = new StyleRunProperties();
            RunFonts runFonts27 = new RunFonts(){ ComplexScript = "Times New Roman" };

            styleRunProperties27.Append(runFonts27);

            style33.Append(styleName33);
            style33.Append(basedOn29);
            style33.Append(uIPriority32);
            style33.Append(rsid75);
            style33.Append(styleRunProperties27);

            Style style34 = new Style(){ Type = StyleValues.Character, StyleId = "crayon-o", CustomStyle = true };
            StyleName styleName34 = new StyleName(){ Val = "crayon-o" };
            BasedOn basedOn30 = new BasedOn(){ Val = "DefaultParagraphFont" };
            UIPriority uIPriority33 = new UIPriority(){ Val = 99 };
            Rsid rsid76 = new Rsid(){ Val = "00FA7479" };

            StyleRunProperties styleRunProperties28 = new StyleRunProperties();
            RunFonts runFonts28 = new RunFonts(){ ComplexScript = "Times New Roman" };

            styleRunProperties28.Append(runFonts28);

            style34.Append(styleName34);
            style34.Append(basedOn30);
            style34.Append(uIPriority33);
            style34.Append(rsid76);
            style34.Append(styleRunProperties28);

            Style style35 = new Style(){ Type = StyleValues.Character, StyleId = "crayon-cn", CustomStyle = true };
            StyleName styleName35 = new StyleName(){ Val = "crayon-cn" };
            BasedOn basedOn31 = new BasedOn(){ Val = "DefaultParagraphFont" };
            UIPriority uIPriority34 = new UIPriority(){ Val = 99 };
            Rsid rsid77 = new Rsid(){ Val = "00FA7479" };

            StyleRunProperties styleRunProperties29 = new StyleRunProperties();
            RunFonts runFonts29 = new RunFonts(){ ComplexScript = "Times New Roman" };

            styleRunProperties29.Append(runFonts29);

            style35.Append(styleName35);
            style35.Append(basedOn31);
            style35.Append(uIPriority34);
            style35.Append(rsid77);
            style35.Append(styleRunProperties29);

            styles1.Append(docDefaults1);
            styles1.Append(latentStyles1);
            styles1.Append(style1);
            styles1.Append(style2);
            styles1.Append(style3);
            styles1.Append(style4);
            styles1.Append(style5);
            styles1.Append(style6);
            styles1.Append(style7);
            styles1.Append(style8);
            styles1.Append(style9);
            styles1.Append(style10);
            styles1.Append(style11);
            styles1.Append(style12);
            styles1.Append(style13);
            styles1.Append(style14);
            styles1.Append(style15);
            styles1.Append(style16);
            styles1.Append(style17);
            styles1.Append(style18);
            styles1.Append(style19);
            styles1.Append(style20);
            styles1.Append(style21);
            styles1.Append(style22);
            styles1.Append(style23);
            styles1.Append(style24);
            styles1.Append(style25);
            styles1.Append(style26);
            styles1.Append(style27);
            styles1.Append(style28);
            styles1.Append(style29);
            styles1.Append(style30);
            styles1.Append(style31);
            styles1.Append(style32);
            styles1.Append(style33);
            styles1.Append(style34);
            styles1.Append(style35);

            styleDefinitionsPart1.Styles = styles1;
        }

        // Generates content of numberingDefinitionsPart1.
        private void GenerateNumberingDefinitionsPart1Content(NumberingDefinitionsPart numberingDefinitionsPart1)
        {
            Numbering numbering1 = new Numbering();
            numbering1.AddNamespaceDeclaration("ve", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            numbering1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            numbering1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            numbering1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            numbering1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            numbering1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            numbering1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            numbering1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            numbering1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");

            AbstractNum abstractNum1 = new AbstractNum(){ AbstractNumberId = 0 };
            Nsid nsid1 = new Nsid(){ Val = "30E032F8" };
            MultiLevelType multiLevelType1 = new MultiLevelType(){ Val = MultiLevelValues.HybridMultilevel };
            TemplateCode templateCode1 = new TemplateCode(){ Val = "F6F6E334" };

            Level level1 = new Level(){ LevelIndex = 0, TemplateCode = "0419000F" };
            StartNumberingValue startNumberingValue1 = new StartNumberingValue(){ Val = 1 };
            NumberingFormat numberingFormat1 = new NumberingFormat(){ Val = NumberFormatValues.Decimal };
            LevelText levelText1 = new LevelText(){ Val = "%1." };
            LevelJustification levelJustification1 = new LevelJustification(){ Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties1 = new PreviousParagraphProperties();
            Indentation indentation2 = new Indentation(){ Left = "720", Hanging = "360" };

            previousParagraphProperties1.Append(indentation2);

            NumberingSymbolRunProperties numberingSymbolRunProperties1 = new NumberingSymbolRunProperties();
            RunFonts runFonts30 = new RunFonts(){ Hint = FontTypeHintValues.Default, ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties1.Append(runFonts30);

            level1.Append(startNumberingValue1);
            level1.Append(numberingFormat1);
            level1.Append(levelText1);
            level1.Append(levelJustification1);
            level1.Append(previousParagraphProperties1);
            level1.Append(numberingSymbolRunProperties1);

            Level level2 = new Level(){ LevelIndex = 1, TemplateCode = "04190019", Tentative = true };
            StartNumberingValue startNumberingValue2 = new StartNumberingValue(){ Val = 1 };
            NumberingFormat numberingFormat2 = new NumberingFormat(){ Val = NumberFormatValues.LowerLetter };
            LevelText levelText2 = new LevelText(){ Val = "%2." };
            LevelJustification levelJustification2 = new LevelJustification(){ Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties2 = new PreviousParagraphProperties();
            Indentation indentation3 = new Indentation(){ Left = "1440", Hanging = "360" };

            previousParagraphProperties2.Append(indentation3);

            NumberingSymbolRunProperties numberingSymbolRunProperties2 = new NumberingSymbolRunProperties();
            RunFonts runFonts31 = new RunFonts(){ ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties2.Append(runFonts31);

            level2.Append(startNumberingValue2);
            level2.Append(numberingFormat2);
            level2.Append(levelText2);
            level2.Append(levelJustification2);
            level2.Append(previousParagraphProperties2);
            level2.Append(numberingSymbolRunProperties2);

            Level level3 = new Level(){ LevelIndex = 2, TemplateCode = "0419001B", Tentative = true };
            StartNumberingValue startNumberingValue3 = new StartNumberingValue(){ Val = 1 };
            NumberingFormat numberingFormat3 = new NumberingFormat(){ Val = NumberFormatValues.LowerRoman };
            LevelText levelText3 = new LevelText(){ Val = "%3." };
            LevelJustification levelJustification3 = new LevelJustification(){ Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties3 = new PreviousParagraphProperties();
            Indentation indentation4 = new Indentation(){ Left = "2160", Hanging = "180" };

            previousParagraphProperties3.Append(indentation4);

            NumberingSymbolRunProperties numberingSymbolRunProperties3 = new NumberingSymbolRunProperties();
            RunFonts runFonts32 = new RunFonts(){ ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties3.Append(runFonts32);

            level3.Append(startNumberingValue3);
            level3.Append(numberingFormat3);
            level3.Append(levelText3);
            level3.Append(levelJustification3);
            level3.Append(previousParagraphProperties3);
            level3.Append(numberingSymbolRunProperties3);

            Level level4 = new Level(){ LevelIndex = 3, TemplateCode = "0419000F", Tentative = true };
            StartNumberingValue startNumberingValue4 = new StartNumberingValue(){ Val = 1 };
            NumberingFormat numberingFormat4 = new NumberingFormat(){ Val = NumberFormatValues.Decimal };
            LevelText levelText4 = new LevelText(){ Val = "%4." };
            LevelJustification levelJustification4 = new LevelJustification(){ Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties4 = new PreviousParagraphProperties();
            Indentation indentation5 = new Indentation(){ Left = "2880", Hanging = "360" };

            previousParagraphProperties4.Append(indentation5);

            NumberingSymbolRunProperties numberingSymbolRunProperties4 = new NumberingSymbolRunProperties();
            RunFonts runFonts33 = new RunFonts(){ ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties4.Append(runFonts33);

            level4.Append(startNumberingValue4);
            level4.Append(numberingFormat4);
            level4.Append(levelText4);
            level4.Append(levelJustification4);
            level4.Append(previousParagraphProperties4);
            level4.Append(numberingSymbolRunProperties4);

            Level level5 = new Level(){ LevelIndex = 4, TemplateCode = "04190019", Tentative = true };
            StartNumberingValue startNumberingValue5 = new StartNumberingValue(){ Val = 1 };
            NumberingFormat numberingFormat5 = new NumberingFormat(){ Val = NumberFormatValues.LowerLetter };
            LevelText levelText5 = new LevelText(){ Val = "%5." };
            LevelJustification levelJustification5 = new LevelJustification(){ Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties5 = new PreviousParagraphProperties();
            Indentation indentation6 = new Indentation(){ Left = "3600", Hanging = "360" };

            previousParagraphProperties5.Append(indentation6);

            NumberingSymbolRunProperties numberingSymbolRunProperties5 = new NumberingSymbolRunProperties();
            RunFonts runFonts34 = new RunFonts(){ ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties5.Append(runFonts34);

            level5.Append(startNumberingValue5);
            level5.Append(numberingFormat5);
            level5.Append(levelText5);
            level5.Append(levelJustification5);
            level5.Append(previousParagraphProperties5);
            level5.Append(numberingSymbolRunProperties5);

            Level level6 = new Level(){ LevelIndex = 5, TemplateCode = "0419001B", Tentative = true };
            StartNumberingValue startNumberingValue6 = new StartNumberingValue(){ Val = 1 };
            NumberingFormat numberingFormat6 = new NumberingFormat(){ Val = NumberFormatValues.LowerRoman };
            LevelText levelText6 = new LevelText(){ Val = "%6." };
            LevelJustification levelJustification6 = new LevelJustification(){ Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties6 = new PreviousParagraphProperties();
            Indentation indentation7 = new Indentation(){ Left = "4320", Hanging = "180" };

            previousParagraphProperties6.Append(indentation7);

            NumberingSymbolRunProperties numberingSymbolRunProperties6 = new NumberingSymbolRunProperties();
            RunFonts runFonts35 = new RunFonts(){ ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties6.Append(runFonts35);

            level6.Append(startNumberingValue6);
            level6.Append(numberingFormat6);
            level6.Append(levelText6);
            level6.Append(levelJustification6);
            level6.Append(previousParagraphProperties6);
            level6.Append(numberingSymbolRunProperties6);

            Level level7 = new Level(){ LevelIndex = 6, TemplateCode = "0419000F", Tentative = true };
            StartNumberingValue startNumberingValue7 = new StartNumberingValue(){ Val = 1 };
            NumberingFormat numberingFormat7 = new NumberingFormat(){ Val = NumberFormatValues.Decimal };
            LevelText levelText7 = new LevelText(){ Val = "%7." };
            LevelJustification levelJustification7 = new LevelJustification(){ Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties7 = new PreviousParagraphProperties();
            Indentation indentation8 = new Indentation(){ Left = "5040", Hanging = "360" };

            previousParagraphProperties7.Append(indentation8);

            NumberingSymbolRunProperties numberingSymbolRunProperties7 = new NumberingSymbolRunProperties();
            RunFonts runFonts36 = new RunFonts(){ ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties7.Append(runFonts36);

            level7.Append(startNumberingValue7);
            level7.Append(numberingFormat7);
            level7.Append(levelText7);
            level7.Append(levelJustification7);
            level7.Append(previousParagraphProperties7);
            level7.Append(numberingSymbolRunProperties7);

            Level level8 = new Level(){ LevelIndex = 7, TemplateCode = "04190019", Tentative = true };
            StartNumberingValue startNumberingValue8 = new StartNumberingValue(){ Val = 1 };
            NumberingFormat numberingFormat8 = new NumberingFormat(){ Val = NumberFormatValues.LowerLetter };
            LevelText levelText8 = new LevelText(){ Val = "%8." };
            LevelJustification levelJustification8 = new LevelJustification(){ Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties8 = new PreviousParagraphProperties();
            Indentation indentation9 = new Indentation(){ Left = "5760", Hanging = "360" };

            previousParagraphProperties8.Append(indentation9);

            NumberingSymbolRunProperties numberingSymbolRunProperties8 = new NumberingSymbolRunProperties();
            RunFonts runFonts37 = new RunFonts(){ ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties8.Append(runFonts37);

            level8.Append(startNumberingValue8);
            level8.Append(numberingFormat8);
            level8.Append(levelText8);
            level8.Append(levelJustification8);
            level8.Append(previousParagraphProperties8);
            level8.Append(numberingSymbolRunProperties8);

            Level level9 = new Level(){ LevelIndex = 8, TemplateCode = "0419001B", Tentative = true };
            StartNumberingValue startNumberingValue9 = new StartNumberingValue(){ Val = 1 };
            NumberingFormat numberingFormat9 = new NumberingFormat(){ Val = NumberFormatValues.LowerRoman };
            LevelText levelText9 = new LevelText(){ Val = "%9." };
            LevelJustification levelJustification9 = new LevelJustification(){ Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties9 = new PreviousParagraphProperties();
            Indentation indentation10 = new Indentation(){ Left = "6480", Hanging = "180" };

            previousParagraphProperties9.Append(indentation10);

            NumberingSymbolRunProperties numberingSymbolRunProperties9 = new NumberingSymbolRunProperties();
            RunFonts runFonts38 = new RunFonts(){ ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties9.Append(runFonts38);

            level9.Append(startNumberingValue9);
            level9.Append(numberingFormat9);
            level9.Append(levelText9);
            level9.Append(levelJustification9);
            level9.Append(previousParagraphProperties9);
            level9.Append(numberingSymbolRunProperties9);

            abstractNum1.Append(nsid1);
            abstractNum1.Append(multiLevelType1);
            abstractNum1.Append(templateCode1);
            abstractNum1.Append(level1);
            abstractNum1.Append(level2);
            abstractNum1.Append(level3);
            abstractNum1.Append(level4);
            abstractNum1.Append(level5);
            abstractNum1.Append(level6);
            abstractNum1.Append(level7);
            abstractNum1.Append(level8);
            abstractNum1.Append(level9);

            AbstractNum abstractNum2 = new AbstractNum(){ AbstractNumberId = 1 };
            Nsid nsid2 = new Nsid(){ Val = "767369AC" };
            MultiLevelType multiLevelType2 = new MultiLevelType(){ Val = MultiLevelValues.HybridMultilevel };
            TemplateCode templateCode2 = new TemplateCode(){ Val = "F6F6E334" };

            Level level10 = new Level(){ LevelIndex = 0, TemplateCode = "0419000F" };
            StartNumberingValue startNumberingValue10 = new StartNumberingValue(){ Val = 1 };
            NumberingFormat numberingFormat10 = new NumberingFormat(){ Val = NumberFormatValues.Decimal };
            LevelText levelText10 = new LevelText(){ Val = "%1." };
            LevelJustification levelJustification10 = new LevelJustification(){ Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties10 = new PreviousParagraphProperties();
            Indentation indentation11 = new Indentation(){ Left = "720", Hanging = "360" };

            previousParagraphProperties10.Append(indentation11);

            NumberingSymbolRunProperties numberingSymbolRunProperties10 = new NumberingSymbolRunProperties();
            RunFonts runFonts39 = new RunFonts(){ Hint = FontTypeHintValues.Default, ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties10.Append(runFonts39);

            level10.Append(startNumberingValue10);
            level10.Append(numberingFormat10);
            level10.Append(levelText10);
            level10.Append(levelJustification10);
            level10.Append(previousParagraphProperties10);
            level10.Append(numberingSymbolRunProperties10);

            Level level11 = new Level(){ LevelIndex = 1, TemplateCode = "04190019", Tentative = true };
            StartNumberingValue startNumberingValue11 = new StartNumberingValue(){ Val = 1 };
            NumberingFormat numberingFormat11 = new NumberingFormat(){ Val = NumberFormatValues.LowerLetter };
            LevelText levelText11 = new LevelText(){ Val = "%2." };
            LevelJustification levelJustification11 = new LevelJustification(){ Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties11 = new PreviousParagraphProperties();
            Indentation indentation12 = new Indentation(){ Left = "1440", Hanging = "360" };

            previousParagraphProperties11.Append(indentation12);

            NumberingSymbolRunProperties numberingSymbolRunProperties11 = new NumberingSymbolRunProperties();
            RunFonts runFonts40 = new RunFonts(){ ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties11.Append(runFonts40);

            level11.Append(startNumberingValue11);
            level11.Append(numberingFormat11);
            level11.Append(levelText11);
            level11.Append(levelJustification11);
            level11.Append(previousParagraphProperties11);
            level11.Append(numberingSymbolRunProperties11);

            Level level12 = new Level(){ LevelIndex = 2, TemplateCode = "0419001B", Tentative = true };
            StartNumberingValue startNumberingValue12 = new StartNumberingValue(){ Val = 1 };
            NumberingFormat numberingFormat12 = new NumberingFormat(){ Val = NumberFormatValues.LowerRoman };
            LevelText levelText12 = new LevelText(){ Val = "%3." };
            LevelJustification levelJustification12 = new LevelJustification(){ Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties12 = new PreviousParagraphProperties();
            Indentation indentation13 = new Indentation(){ Left = "2160", Hanging = "180" };

            previousParagraphProperties12.Append(indentation13);

            NumberingSymbolRunProperties numberingSymbolRunProperties12 = new NumberingSymbolRunProperties();
            RunFonts runFonts41 = new RunFonts(){ ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties12.Append(runFonts41);

            level12.Append(startNumberingValue12);
            level12.Append(numberingFormat12);
            level12.Append(levelText12);
            level12.Append(levelJustification12);
            level12.Append(previousParagraphProperties12);
            level12.Append(numberingSymbolRunProperties12);

            Level level13 = new Level(){ LevelIndex = 3, TemplateCode = "0419000F", Tentative = true };
            StartNumberingValue startNumberingValue13 = new StartNumberingValue(){ Val = 1 };
            NumberingFormat numberingFormat13 = new NumberingFormat(){ Val = NumberFormatValues.Decimal };
            LevelText levelText13 = new LevelText(){ Val = "%4." };
            LevelJustification levelJustification13 = new LevelJustification(){ Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties13 = new PreviousParagraphProperties();
            Indentation indentation14 = new Indentation(){ Left = "2880", Hanging = "360" };

            previousParagraphProperties13.Append(indentation14);

            NumberingSymbolRunProperties numberingSymbolRunProperties13 = new NumberingSymbolRunProperties();
            RunFonts runFonts42 = new RunFonts(){ ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties13.Append(runFonts42);

            level13.Append(startNumberingValue13);
            level13.Append(numberingFormat13);
            level13.Append(levelText13);
            level13.Append(levelJustification13);
            level13.Append(previousParagraphProperties13);
            level13.Append(numberingSymbolRunProperties13);

            Level level14 = new Level(){ LevelIndex = 4, TemplateCode = "04190019", Tentative = true };
            StartNumberingValue startNumberingValue14 = new StartNumberingValue(){ Val = 1 };
            NumberingFormat numberingFormat14 = new NumberingFormat(){ Val = NumberFormatValues.LowerLetter };
            LevelText levelText14 = new LevelText(){ Val = "%5." };
            LevelJustification levelJustification14 = new LevelJustification(){ Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties14 = new PreviousParagraphProperties();
            Indentation indentation15 = new Indentation(){ Left = "3600", Hanging = "360" };

            previousParagraphProperties14.Append(indentation15);

            NumberingSymbolRunProperties numberingSymbolRunProperties14 = new NumberingSymbolRunProperties();
            RunFonts runFonts43 = new RunFonts(){ ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties14.Append(runFonts43);

            level14.Append(startNumberingValue14);
            level14.Append(numberingFormat14);
            level14.Append(levelText14);
            level14.Append(levelJustification14);
            level14.Append(previousParagraphProperties14);
            level14.Append(numberingSymbolRunProperties14);

            Level level15 = new Level(){ LevelIndex = 5, TemplateCode = "0419001B", Tentative = true };
            StartNumberingValue startNumberingValue15 = new StartNumberingValue(){ Val = 1 };
            NumberingFormat numberingFormat15 = new NumberingFormat(){ Val = NumberFormatValues.LowerRoman };
            LevelText levelText15 = new LevelText(){ Val = "%6." };
            LevelJustification levelJustification15 = new LevelJustification(){ Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties15 = new PreviousParagraphProperties();
            Indentation indentation16 = new Indentation(){ Left = "4320", Hanging = "180" };

            previousParagraphProperties15.Append(indentation16);

            NumberingSymbolRunProperties numberingSymbolRunProperties15 = new NumberingSymbolRunProperties();
            RunFonts runFonts44 = new RunFonts(){ ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties15.Append(runFonts44);

            level15.Append(startNumberingValue15);
            level15.Append(numberingFormat15);
            level15.Append(levelText15);
            level15.Append(levelJustification15);
            level15.Append(previousParagraphProperties15);
            level15.Append(numberingSymbolRunProperties15);

            Level level16 = new Level(){ LevelIndex = 6, TemplateCode = "0419000F", Tentative = true };
            StartNumberingValue startNumberingValue16 = new StartNumberingValue(){ Val = 1 };
            NumberingFormat numberingFormat16 = new NumberingFormat(){ Val = NumberFormatValues.Decimal };
            LevelText levelText16 = new LevelText(){ Val = "%7." };
            LevelJustification levelJustification16 = new LevelJustification(){ Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties16 = new PreviousParagraphProperties();
            Indentation indentation17 = new Indentation(){ Left = "5040", Hanging = "360" };

            previousParagraphProperties16.Append(indentation17);

            NumberingSymbolRunProperties numberingSymbolRunProperties16 = new NumberingSymbolRunProperties();
            RunFonts runFonts45 = new RunFonts(){ ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties16.Append(runFonts45);

            level16.Append(startNumberingValue16);
            level16.Append(numberingFormat16);
            level16.Append(levelText16);
            level16.Append(levelJustification16);
            level16.Append(previousParagraphProperties16);
            level16.Append(numberingSymbolRunProperties16);

            Level level17 = new Level(){ LevelIndex = 7, TemplateCode = "04190019", Tentative = true };
            StartNumberingValue startNumberingValue17 = new StartNumberingValue(){ Val = 1 };
            NumberingFormat numberingFormat17 = new NumberingFormat(){ Val = NumberFormatValues.LowerLetter };
            LevelText levelText17 = new LevelText(){ Val = "%8." };
            LevelJustification levelJustification17 = new LevelJustification(){ Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties17 = new PreviousParagraphProperties();
            Indentation indentation18 = new Indentation(){ Left = "5760", Hanging = "360" };

            previousParagraphProperties17.Append(indentation18);

            NumberingSymbolRunProperties numberingSymbolRunProperties17 = new NumberingSymbolRunProperties();
            RunFonts runFonts46 = new RunFonts(){ ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties17.Append(runFonts46);

            level17.Append(startNumberingValue17);
            level17.Append(numberingFormat17);
            level17.Append(levelText17);
            level17.Append(levelJustification17);
            level17.Append(previousParagraphProperties17);
            level17.Append(numberingSymbolRunProperties17);

            Level level18 = new Level(){ LevelIndex = 8, TemplateCode = "0419001B", Tentative = true };
            StartNumberingValue startNumberingValue18 = new StartNumberingValue(){ Val = 1 };
            NumberingFormat numberingFormat18 = new NumberingFormat(){ Val = NumberFormatValues.LowerRoman };
            LevelText levelText18 = new LevelText(){ Val = "%9." };
            LevelJustification levelJustification18 = new LevelJustification(){ Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties18 = new PreviousParagraphProperties();
            Indentation indentation19 = new Indentation(){ Left = "6480", Hanging = "180" };

            previousParagraphProperties18.Append(indentation19);

            NumberingSymbolRunProperties numberingSymbolRunProperties18 = new NumberingSymbolRunProperties();
            RunFonts runFonts47 = new RunFonts(){ ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties18.Append(runFonts47);

            level18.Append(startNumberingValue18);
            level18.Append(numberingFormat18);
            level18.Append(levelText18);
            level18.Append(levelJustification18);
            level18.Append(previousParagraphProperties18);
            level18.Append(numberingSymbolRunProperties18);

            abstractNum2.Append(nsid2);
            abstractNum2.Append(multiLevelType2);
            abstractNum2.Append(templateCode2);
            abstractNum2.Append(level10);
            abstractNum2.Append(level11);
            abstractNum2.Append(level12);
            abstractNum2.Append(level13);
            abstractNum2.Append(level14);
            abstractNum2.Append(level15);
            abstractNum2.Append(level16);
            abstractNum2.Append(level17);
            abstractNum2.Append(level18);

            NumberingInstance numberingInstance1 = new NumberingInstance(){ NumberID = 1 };
            AbstractNumId abstractNumId1 = new AbstractNumId(){ Val = 0 };

            numberingInstance1.Append(abstractNumId1);

            NumberingInstance numberingInstance2 = new NumberingInstance(){ NumberID = 2 };
            AbstractNumId abstractNumId2 = new AbstractNumId(){ Val = 1 };

            numberingInstance2.Append(abstractNumId2);

            numbering1.Append(abstractNum1);
            numbering1.Append(abstractNum2);
            numbering1.Append(numberingInstance1);
            numbering1.Append(numberingInstance2);

            numberingDefinitionsPart1.Numbering = numbering1;
        }

        // Generates content of endnotesPart1.
        private void GenerateEndnotesPart1Content(EndnotesPart endnotesPart1)
        {
            Endnotes endnotes1 = new Endnotes();
            endnotes1.AddNamespaceDeclaration("ve", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            endnotes1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            endnotes1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            endnotes1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            endnotes1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            endnotes1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            endnotes1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            endnotes1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            endnotes1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");

            Endnote endnote1 = new Endnote(){ Type = FootnoteEndnoteValues.Separator, Id = -1 };

            Paragraph paragraph2 = new Paragraph(){ RsidParagraphAddition = "00C3696F", RsidParagraphProperties = "007D2F02", RsidRunAdditionDefault = "00C3696F" };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines7 = new SpacingBetweenLines(){ After = "0" };

            paragraphProperties2.Append(spacingBetweenLines7);

            Run run2 = new Run();
            SeparatorMark separatorMark1 = new SeparatorMark();

            run2.Append(separatorMark1);

            paragraph2.Append(paragraphProperties2);
            paragraph2.Append(run2);

            endnote1.Append(paragraph2);

            Endnote endnote2 = new Endnote(){ Type = FootnoteEndnoteValues.ContinuationSeparator, Id = 0 };

            Paragraph paragraph3 = new Paragraph(){ RsidParagraphAddition = "00C3696F", RsidParagraphProperties = "007D2F02", RsidRunAdditionDefault = "00C3696F" };

            ParagraphProperties paragraphProperties3 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines8 = new SpacingBetweenLines(){ After = "0" };

            paragraphProperties3.Append(spacingBetweenLines8);

            Run run3 = new Run();
            ContinuationSeparatorMark continuationSeparatorMark1 = new ContinuationSeparatorMark();

            run3.Append(continuationSeparatorMark1);

            paragraph3.Append(paragraphProperties3);
            paragraph3.Append(run3);

            endnote2.Append(paragraph3);

            endnotes1.Append(endnote1);
            endnotes1.Append(endnote2);

            endnotesPart1.Endnotes = endnotes1;
        }

        // Generates content of footnotesPart1.
        private void GenerateFootnotesPart1Content(FootnotesPart footnotesPart1)
        {
            Footnotes footnotes1 = new Footnotes();
            footnotes1.AddNamespaceDeclaration("ve", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            footnotes1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            footnotes1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            footnotes1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            footnotes1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            footnotes1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            footnotes1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            footnotes1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            footnotes1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");

            Footnote footnote1 = new Footnote(){ Type = FootnoteEndnoteValues.Separator, Id = -1 };

            Paragraph paragraph4 = new Paragraph(){ RsidParagraphAddition = "00C3696F", RsidParagraphProperties = "007D2F02", RsidRunAdditionDefault = "00C3696F" };

            ParagraphProperties paragraphProperties4 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines9 = new SpacingBetweenLines(){ After = "0" };

            paragraphProperties4.Append(spacingBetweenLines9);

            Run run4 = new Run();
            SeparatorMark separatorMark2 = new SeparatorMark();

            run4.Append(separatorMark2);

            paragraph4.Append(paragraphProperties4);
            paragraph4.Append(run4);

            footnote1.Append(paragraph4);

            Footnote footnote2 = new Footnote(){ Type = FootnoteEndnoteValues.ContinuationSeparator, Id = 0 };

            Paragraph paragraph5 = new Paragraph(){ RsidParagraphAddition = "00C3696F", RsidParagraphProperties = "007D2F02", RsidRunAdditionDefault = "00C3696F" };

            ParagraphProperties paragraphProperties5 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines10 = new SpacingBetweenLines(){ After = "0" };

            paragraphProperties5.Append(spacingBetweenLines10);

            Run run5 = new Run();
            ContinuationSeparatorMark continuationSeparatorMark2 = new ContinuationSeparatorMark();

            run5.Append(continuationSeparatorMark2);

            paragraph5.Append(paragraphProperties5);
            paragraph5.Append(run5);

            footnote2.Append(paragraph5);

            footnotes1.Append(footnote1);
            footnotes1.Append(footnote2);

            footnotesPart1.Footnotes = footnotes1;
        }

        // Generates content of webSettingsPart1.
        private void GenerateWebSettingsPart1Content(WebSettingsPart webSettingsPart1)
        {
            WebSettings webSettings1 = new WebSettings();
            webSettings1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            webSettings1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            Divs divs1 = new Divs();

            Div div1 = new Div(){ Id = "2039310386" };
            LeftMarginDiv leftMarginDiv1 = new LeftMarginDiv(){ Val = "0" };
            RightMarginDiv rightMarginDiv1 = new RightMarginDiv(){ Val = "0" };
            TopMarginDiv topMarginDiv1 = new TopMarginDiv(){ Val = "0" };
            BottomMarginDiv bottomMarginDiv1 = new BottomMarginDiv(){ Val = "0" };

            DivBorder divBorder1 = new DivBorder();
            TopBorder topBorder2 = new TopBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder2 = new LeftBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder2 = new BottomBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder2 = new RightBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            divBorder1.Append(topBorder2);
            divBorder1.Append(leftBorder2);
            divBorder1.Append(bottomBorder2);
            divBorder1.Append(rightBorder2);

            div1.Append(leftMarginDiv1);
            div1.Append(rightMarginDiv1);
            div1.Append(topMarginDiv1);
            div1.Append(bottomMarginDiv1);
            div1.Append(divBorder1);

            Div div2 = new Div(){ Id = "2039310387" };
            LeftMarginDiv leftMarginDiv2 = new LeftMarginDiv(){ Val = "0" };
            RightMarginDiv rightMarginDiv2 = new RightMarginDiv(){ Val = "0" };
            TopMarginDiv topMarginDiv2 = new TopMarginDiv(){ Val = "0" };
            BottomMarginDiv bottomMarginDiv2 = new BottomMarginDiv(){ Val = "0" };

            DivBorder divBorder2 = new DivBorder();
            TopBorder topBorder3 = new TopBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder3 = new LeftBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder3 = new BottomBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder3 = new RightBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            divBorder2.Append(topBorder3);
            divBorder2.Append(leftBorder3);
            divBorder2.Append(bottomBorder3);
            divBorder2.Append(rightBorder3);

            div2.Append(leftMarginDiv2);
            div2.Append(rightMarginDiv2);
            div2.Append(topMarginDiv2);
            div2.Append(bottomMarginDiv2);
            div2.Append(divBorder2);

            Div div3 = new Div(){ Id = "2039310388" };
            LeftMarginDiv leftMarginDiv3 = new LeftMarginDiv(){ Val = "0" };
            RightMarginDiv rightMarginDiv3 = new RightMarginDiv(){ Val = "0" };
            TopMarginDiv topMarginDiv3 = new TopMarginDiv(){ Val = "0" };
            BottomMarginDiv bottomMarginDiv3 = new BottomMarginDiv(){ Val = "0" };

            DivBorder divBorder3 = new DivBorder();
            TopBorder topBorder4 = new TopBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder4 = new LeftBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder4 = new BottomBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder4 = new RightBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            divBorder3.Append(topBorder4);
            divBorder3.Append(leftBorder4);
            divBorder3.Append(bottomBorder4);
            divBorder3.Append(rightBorder4);

            div3.Append(leftMarginDiv3);
            div3.Append(rightMarginDiv3);
            div3.Append(topMarginDiv3);
            div3.Append(bottomMarginDiv3);
            div3.Append(divBorder3);

            Div div4 = new Div(){ Id = "2039310389" };
            LeftMarginDiv leftMarginDiv4 = new LeftMarginDiv(){ Val = "0" };
            RightMarginDiv rightMarginDiv4 = new RightMarginDiv(){ Val = "0" };
            TopMarginDiv topMarginDiv4 = new TopMarginDiv(){ Val = "0" };
            BottomMarginDiv bottomMarginDiv4 = new BottomMarginDiv(){ Val = "0" };

            DivBorder divBorder4 = new DivBorder();
            TopBorder topBorder5 = new TopBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder5 = new LeftBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder5 = new BottomBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder5 = new RightBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            divBorder4.Append(topBorder5);
            divBorder4.Append(leftBorder5);
            divBorder4.Append(bottomBorder5);
            divBorder4.Append(rightBorder5);

            div4.Append(leftMarginDiv4);
            div4.Append(rightMarginDiv4);
            div4.Append(topMarginDiv4);
            div4.Append(bottomMarginDiv4);
            div4.Append(divBorder4);

            Div div5 = new Div(){ Id = "2039310390" };
            LeftMarginDiv leftMarginDiv5 = new LeftMarginDiv(){ Val = "0" };
            RightMarginDiv rightMarginDiv5 = new RightMarginDiv(){ Val = "0" };
            TopMarginDiv topMarginDiv5 = new TopMarginDiv(){ Val = "0" };
            BottomMarginDiv bottomMarginDiv5 = new BottomMarginDiv(){ Val = "0" };

            DivBorder divBorder5 = new DivBorder();
            TopBorder topBorder6 = new TopBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder6 = new LeftBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder6 = new BottomBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder6 = new RightBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            divBorder5.Append(topBorder6);
            divBorder5.Append(leftBorder6);
            divBorder5.Append(bottomBorder6);
            divBorder5.Append(rightBorder6);

            div5.Append(leftMarginDiv5);
            div5.Append(rightMarginDiv5);
            div5.Append(topMarginDiv5);
            div5.Append(bottomMarginDiv5);
            div5.Append(divBorder5);

            Div div6 = new Div(){ Id = "2039310391" };
            LeftMarginDiv leftMarginDiv6 = new LeftMarginDiv(){ Val = "0" };
            RightMarginDiv rightMarginDiv6 = new RightMarginDiv(){ Val = "0" };
            TopMarginDiv topMarginDiv6 = new TopMarginDiv(){ Val = "0" };
            BottomMarginDiv bottomMarginDiv6 = new BottomMarginDiv(){ Val = "0" };

            DivBorder divBorder6 = new DivBorder();
            TopBorder topBorder7 = new TopBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder7 = new LeftBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder7 = new BottomBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder7 = new RightBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            divBorder6.Append(topBorder7);
            divBorder6.Append(leftBorder7);
            divBorder6.Append(bottomBorder7);
            divBorder6.Append(rightBorder7);

            div6.Append(leftMarginDiv6);
            div6.Append(rightMarginDiv6);
            div6.Append(topMarginDiv6);
            div6.Append(bottomMarginDiv6);
            div6.Append(divBorder6);

            Div div7 = new Div(){ Id = "2039310392" };
            LeftMarginDiv leftMarginDiv7 = new LeftMarginDiv(){ Val = "0" };
            RightMarginDiv rightMarginDiv7 = new RightMarginDiv(){ Val = "0" };
            TopMarginDiv topMarginDiv7 = new TopMarginDiv(){ Val = "0" };
            BottomMarginDiv bottomMarginDiv7 = new BottomMarginDiv(){ Val = "0" };

            DivBorder divBorder7 = new DivBorder();
            TopBorder topBorder8 = new TopBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder8 = new LeftBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder8 = new BottomBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder8 = new RightBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            divBorder7.Append(topBorder8);
            divBorder7.Append(leftBorder8);
            divBorder7.Append(bottomBorder8);
            divBorder7.Append(rightBorder8);

            div7.Append(leftMarginDiv7);
            div7.Append(rightMarginDiv7);
            div7.Append(topMarginDiv7);
            div7.Append(bottomMarginDiv7);
            div7.Append(divBorder7);

            Div div8 = new Div(){ Id = "2039310393" };
            LeftMarginDiv leftMarginDiv8 = new LeftMarginDiv(){ Val = "0" };
            RightMarginDiv rightMarginDiv8 = new RightMarginDiv(){ Val = "0" };
            TopMarginDiv topMarginDiv8 = new TopMarginDiv(){ Val = "0" };
            BottomMarginDiv bottomMarginDiv8 = new BottomMarginDiv(){ Val = "0" };

            DivBorder divBorder8 = new DivBorder();
            TopBorder topBorder9 = new TopBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder9 = new LeftBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder9 = new BottomBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder9 = new RightBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            divBorder8.Append(topBorder9);
            divBorder8.Append(leftBorder9);
            divBorder8.Append(bottomBorder9);
            divBorder8.Append(rightBorder9);

            div8.Append(leftMarginDiv8);
            div8.Append(rightMarginDiv8);
            div8.Append(topMarginDiv8);
            div8.Append(bottomMarginDiv8);
            div8.Append(divBorder8);

            Div div9 = new Div(){ Id = "2039310394" };
            LeftMarginDiv leftMarginDiv9 = new LeftMarginDiv(){ Val = "0" };
            RightMarginDiv rightMarginDiv9 = new RightMarginDiv(){ Val = "0" };
            TopMarginDiv topMarginDiv9 = new TopMarginDiv(){ Val = "0" };
            BottomMarginDiv bottomMarginDiv9 = new BottomMarginDiv(){ Val = "0" };

            DivBorder divBorder9 = new DivBorder();
            TopBorder topBorder10 = new TopBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder10 = new LeftBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder10 = new BottomBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder10 = new RightBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            divBorder9.Append(topBorder10);
            divBorder9.Append(leftBorder10);
            divBorder9.Append(bottomBorder10);
            divBorder9.Append(rightBorder10);

            div9.Append(leftMarginDiv9);
            div9.Append(rightMarginDiv9);
            div9.Append(topMarginDiv9);
            div9.Append(bottomMarginDiv9);
            div9.Append(divBorder9);

            Div div10 = new Div(){ Id = "2039310395" };
            LeftMarginDiv leftMarginDiv10 = new LeftMarginDiv(){ Val = "0" };
            RightMarginDiv rightMarginDiv10 = new RightMarginDiv(){ Val = "0" };
            TopMarginDiv topMarginDiv10 = new TopMarginDiv(){ Val = "0" };
            BottomMarginDiv bottomMarginDiv10 = new BottomMarginDiv(){ Val = "0" };

            DivBorder divBorder10 = new DivBorder();
            TopBorder topBorder11 = new TopBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder11 = new LeftBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder11 = new BottomBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder11 = new RightBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            divBorder10.Append(topBorder11);
            divBorder10.Append(leftBorder11);
            divBorder10.Append(bottomBorder11);
            divBorder10.Append(rightBorder11);

            div10.Append(leftMarginDiv10);
            div10.Append(rightMarginDiv10);
            div10.Append(topMarginDiv10);
            div10.Append(bottomMarginDiv10);
            div10.Append(divBorder10);

            Div div11 = new Div(){ Id = "2039310396" };
            LeftMarginDiv leftMarginDiv11 = new LeftMarginDiv(){ Val = "0" };
            RightMarginDiv rightMarginDiv11 = new RightMarginDiv(){ Val = "0" };
            TopMarginDiv topMarginDiv11 = new TopMarginDiv(){ Val = "0" };
            BottomMarginDiv bottomMarginDiv11 = new BottomMarginDiv(){ Val = "0" };

            DivBorder divBorder11 = new DivBorder();
            TopBorder topBorder12 = new TopBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder12 = new LeftBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder12 = new BottomBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder12 = new RightBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            divBorder11.Append(topBorder12);
            divBorder11.Append(leftBorder12);
            divBorder11.Append(bottomBorder12);
            divBorder11.Append(rightBorder12);

            div11.Append(leftMarginDiv11);
            div11.Append(rightMarginDiv11);
            div11.Append(topMarginDiv11);
            div11.Append(bottomMarginDiv11);
            div11.Append(divBorder11);

            Div div12 = new Div(){ Id = "2039310397" };
            LeftMarginDiv leftMarginDiv12 = new LeftMarginDiv(){ Val = "0" };
            RightMarginDiv rightMarginDiv12 = new RightMarginDiv(){ Val = "0" };
            TopMarginDiv topMarginDiv12 = new TopMarginDiv(){ Val = "0" };
            BottomMarginDiv bottomMarginDiv12 = new BottomMarginDiv(){ Val = "0" };

            DivBorder divBorder12 = new DivBorder();
            TopBorder topBorder13 = new TopBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder13 = new LeftBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder13 = new BottomBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder13 = new RightBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            divBorder12.Append(topBorder13);
            divBorder12.Append(leftBorder13);
            divBorder12.Append(bottomBorder13);
            divBorder12.Append(rightBorder13);

            div12.Append(leftMarginDiv12);
            div12.Append(rightMarginDiv12);
            div12.Append(topMarginDiv12);
            div12.Append(bottomMarginDiv12);
            div12.Append(divBorder12);

            Div div13 = new Div(){ Id = "2039310398" };
            LeftMarginDiv leftMarginDiv13 = new LeftMarginDiv(){ Val = "0" };
            RightMarginDiv rightMarginDiv13 = new RightMarginDiv(){ Val = "0" };
            TopMarginDiv topMarginDiv13 = new TopMarginDiv(){ Val = "0" };
            BottomMarginDiv bottomMarginDiv13 = new BottomMarginDiv(){ Val = "0" };

            DivBorder divBorder13 = new DivBorder();
            TopBorder topBorder14 = new TopBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder14 = new LeftBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder14 = new BottomBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder14 = new RightBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            divBorder13.Append(topBorder14);
            divBorder13.Append(leftBorder14);
            divBorder13.Append(bottomBorder14);
            divBorder13.Append(rightBorder14);

            div13.Append(leftMarginDiv13);
            div13.Append(rightMarginDiv13);
            div13.Append(topMarginDiv13);
            div13.Append(bottomMarginDiv13);
            div13.Append(divBorder13);

            Div div14 = new Div(){ Id = "2039310399" };
            LeftMarginDiv leftMarginDiv14 = new LeftMarginDiv(){ Val = "0" };
            RightMarginDiv rightMarginDiv14 = new RightMarginDiv(){ Val = "0" };
            TopMarginDiv topMarginDiv14 = new TopMarginDiv(){ Val = "0" };
            BottomMarginDiv bottomMarginDiv14 = new BottomMarginDiv(){ Val = "0" };

            DivBorder divBorder14 = new DivBorder();
            TopBorder topBorder15 = new TopBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder15 = new LeftBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder15 = new BottomBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder15 = new RightBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            divBorder14.Append(topBorder15);
            divBorder14.Append(leftBorder15);
            divBorder14.Append(bottomBorder15);
            divBorder14.Append(rightBorder15);

            div14.Append(leftMarginDiv14);
            div14.Append(rightMarginDiv14);
            div14.Append(topMarginDiv14);
            div14.Append(bottomMarginDiv14);
            div14.Append(divBorder14);

            Div div15 = new Div(){ Id = "2039310400" };
            LeftMarginDiv leftMarginDiv15 = new LeftMarginDiv(){ Val = "0" };
            RightMarginDiv rightMarginDiv15 = new RightMarginDiv(){ Val = "0" };
            TopMarginDiv topMarginDiv15 = new TopMarginDiv(){ Val = "0" };
            BottomMarginDiv bottomMarginDiv15 = new BottomMarginDiv(){ Val = "0" };

            DivBorder divBorder15 = new DivBorder();
            TopBorder topBorder16 = new TopBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder16 = new LeftBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder16 = new BottomBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder16 = new RightBorder(){ Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            divBorder15.Append(topBorder16);
            divBorder15.Append(leftBorder16);
            divBorder15.Append(bottomBorder16);
            divBorder15.Append(rightBorder16);

            div15.Append(leftMarginDiv15);
            div15.Append(rightMarginDiv15);
            div15.Append(topMarginDiv15);
            div15.Append(bottomMarginDiv15);
            div15.Append(divBorder15);

            divs1.Append(div1);
            divs1.Append(div2);
            divs1.Append(div3);
            divs1.Append(div4);
            divs1.Append(div5);
            divs1.Append(div6);
            divs1.Append(div7);
            divs1.Append(div8);
            divs1.Append(div9);
            divs1.Append(div10);
            divs1.Append(div11);
            divs1.Append(div12);
            divs1.Append(div13);
            divs1.Append(div14);
            divs1.Append(div15);
            OptimizeForBrowser optimizeForBrowser1 = new OptimizeForBrowser();

            webSettings1.Append(divs1);
            webSettings1.Append(optimizeForBrowser1);

            webSettingsPart1.WebSettings = webSettings1;
        }

        private void SetPackageProperties(OpenXmlPackage document)
        {
            document.PackageProperties.Creator = "Полунина Марина";
            document.PackageProperties.Title = "Министерство образования и науки РФ ";
            document.PackageProperties.Subject = "";
            document.PackageProperties.Keywords = "";
            document.PackageProperties.Description = "";
            document.PackageProperties.Revision = "3";
            document.PackageProperties.Created = System.Xml.XmlConvert.ToDateTime("2017-12-10T21:46:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2017-12-20T19:08:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.LastModifiedBy = "Grisha";
        }


    }
}
