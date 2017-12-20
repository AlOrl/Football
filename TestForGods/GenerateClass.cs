using DocumentFormat.OpenXml.Packaging;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using M = DocumentFormat.OpenXml.Math;
using A = DocumentFormat.OpenXml.Drawing;
using TestForGods;

namespace GeneratedCode
{
    public class GeneratedClass:TestViewModel
    {
        
        
        // Creates a WordprocessingDocument.
        public void CreatePackage(string filePath, string message)
        {
            using (WordprocessingDocument package = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
            {
                CreateParts(package, message);
            }
        }

        // Adds child parts and generates content of the specified part.
        private void CreateParts(WordprocessingDocument document, string message)
        {
            ExtendedFilePropertiesPart extendedFilePropertiesPart1 = document.AddNewPart<ExtendedFilePropertiesPart>("rId3");
            GenerateExtendedFilePropertiesPart1Content(extendedFilePropertiesPart1);

            MainDocumentPart mainDocumentPart1 = document.AddMainDocumentPart();
            GenerateMainDocumentPart1Content(mainDocumentPart1, message);

            DocumentSettingsPart documentSettingsPart1 = mainDocumentPart1.AddNewPart<DocumentSettingsPart>("rId3");
            GenerateDocumentSettingsPart1Content(documentSettingsPart1);

            StyleDefinitionsPart styleDefinitionsPart1 = mainDocumentPart1.AddNewPart<StyleDefinitionsPart>("rId2");
            GenerateStyleDefinitionsPart1Content(styleDefinitionsPart1);

            NumberingDefinitionsPart numberingDefinitionsPart1 = mainDocumentPart1.AddNewPart<NumberingDefinitionsPart>("rId1");
            GenerateNumberingDefinitionsPart1Content(numberingDefinitionsPart1);

            ThemePart themePart1 = mainDocumentPart1.AddNewPart<ThemePart>("rId6");
            GenerateThemePart1Content(themePart1);

            FontTablePart fontTablePart1 = mainDocumentPart1.AddNewPart<FontTablePart>("rId5");
            GenerateFontTablePart1Content(fontTablePart1);

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
            totalTime1.Text = "3";
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
            company1.Text = "Microsoft";
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
        private void GenerateMainDocumentPart1Content(MainDocumentPart mainDocumentPart1, string message)
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

            Paragraph paragraph1 = new Paragraph() { RsidParagraphMarkRevision = "00E103E6", RsidParagraphAddition = "00BA3CC1", RsidParagraphProperties = "00CD6E80", RsidRunAdditionDefault = "00BA3CC1" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Line = "259", LineRule = LineSpacingRuleValues.Auto };

            ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
            RunFonts runFonts1 = new RunFonts() { Ascii = "Cambria", HighAnsi = "Cambria" };
            FontSize fontSize1 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "28" };

            paragraphMarkRunProperties1.Append(runFonts1);
            paragraphMarkRunProperties1.Append(fontSize1);
            paragraphMarkRunProperties1.Append(fontSizeComplexScript1);

            paragraphProperties1.Append(spacingBetweenLines1);
            paragraphProperties1.Append(paragraphMarkRunProperties1);

            Run run1 = new Run();

            RunProperties runProperties1 = new RunProperties();
            RunFonts runFonts2 = new RunFonts() { Ascii = "Cambria", HighAnsi = "Cambria" };
            FontSize fontSize2 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "28" };

           

            runProperties1.Append(runFonts2);
            runProperties1.Append(fontSize2);
            runProperties1.Append(fontSizeComplexScript2);
            Text text1 = new Text();
            text1.Text =message;

            run1.Append(runProperties1);
            run1.Append(text1);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);

            SectionProperties sectionProperties1 = new SectionProperties() { RsidRPr = "00E103E6", RsidR = "00BA3CC1", RsidSect = "00756A90" };
            PageSize pageSize1 = new PageSize() { Width = (UInt32Value)11906U, Height = (UInt32Value)16838U };
            PageMargin pageMargin1 = new PageMargin() { Top = 1134, Right = (UInt32Value)851U, Bottom = 1134, Left = (UInt32Value)1418U, Header = (UInt32Value)709U, Footer = (UInt32Value)709U, Gutter = (UInt32Value)0U };
            Columns columns1 = new Columns() { Space = "708" };
            DocGrid docGrid1 = new DocGrid() { LinePitch = 360 };

            sectionProperties1.Append(pageSize1);
            sectionProperties1.Append(pageMargin1);
            sectionProperties1.Append(columns1);
            sectionProperties1.Append(docGrid1);

            body1.Append(paragraph1);
            body1.Append(sectionProperties1);

            document1.Append(body1);

            mainDocumentPart1.Document = document1;
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
            Zoom zoom1 = new Zoom() { Percent = "100" };
            DefaultTabStop defaultTabStop1 = new DefaultTabStop() { Val = 708 };
            DrawingGridHorizontalSpacing drawingGridHorizontalSpacing1 = new DrawingGridHorizontalSpacing() { Val = "181" };
            DrawingGridVerticalSpacing drawingGridVerticalSpacing1 = new DrawingGridVerticalSpacing() { Val = "181" };
            CharacterSpacingControl characterSpacingControl1 = new CharacterSpacingControl() { Val = CharacterSpacingValues.DoNotCompress };

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
            RsidRoot rsidRoot1 = new RsidRoot() { Val = "00D43E79" };
            Rsid rsid1 = new Rsid() { Val = "00006453" };
            Rsid rsid2 = new Rsid() { Val = "00021484" };
            Rsid rsid3 = new Rsid() { Val = "00036752" };
            Rsid rsid4 = new Rsid() { Val = "00076E1D" };
            Rsid rsid5 = new Rsid() { Val = "00086B3F" };
            Rsid rsid6 = new Rsid() { Val = "00096B5E" };
            Rsid rsid7 = new Rsid() { Val = "000E0306" };
            Rsid rsid8 = new Rsid() { Val = "0013270C" };
            Rsid rsid9 = new Rsid() { Val = "002329C2" };
            Rsid rsid10 = new Rsid() { Val = "00282469" };
            Rsid rsid11 = new Rsid() { Val = "00296CAC" };
            Rsid rsid12 = new Rsid() { Val = "002B2FE4" };
            Rsid rsid13 = new Rsid() { Val = "002B6741" };
            Rsid rsid14 = new Rsid() { Val = "002E7749" };
            Rsid rsid15 = new Rsid() { Val = "00325A4A" };
            Rsid rsid16 = new Rsid() { Val = "00395D9C" };
            Rsid rsid17 = new Rsid() { Val = "003F596E" };
            Rsid rsid18 = new Rsid() { Val = "003F7055" };
            Rsid rsid19 = new Rsid() { Val = "00404B8F" };
            Rsid rsid20 = new Rsid() { Val = "004532FB" };
            Rsid rsid21 = new Rsid() { Val = "0048748D" };
            Rsid rsid22 = new Rsid() { Val = "004F1DC1" };
            Rsid rsid23 = new Rsid() { Val = "00504C5D" };
            Rsid rsid24 = new Rsid() { Val = "00544689" };
            Rsid rsid25 = new Rsid() { Val = "00577616" };
            Rsid rsid26 = new Rsid() { Val = "005B2F98" };
            Rsid rsid27 = new Rsid() { Val = "005F07EC" };
            Rsid rsid28 = new Rsid() { Val = "006178FF" };
            Rsid rsid29 = new Rsid() { Val = "0062019B" };
            Rsid rsid30 = new Rsid() { Val = "006550E8" };
            Rsid rsid31 = new Rsid() { Val = "00656FEA" };
            Rsid rsid32 = new Rsid() { Val = "00691515" };
            Rsid rsid33 = new Rsid() { Val = "006B4F88" };
            Rsid rsid34 = new Rsid() { Val = "00717E5B" };
            Rsid rsid35 = new Rsid() { Val = "00741630" };
            Rsid rsid36 = new Rsid() { Val = "00756A90" };
            Rsid rsid37 = new Rsid() { Val = "00756E4A" };
            Rsid rsid38 = new Rsid() { Val = "00787259" };
            Rsid rsid39 = new Rsid() { Val = "007B0123" };
            Rsid rsid40 = new Rsid() { Val = "007C36A0" };
            Rsid rsid41 = new Rsid() { Val = "007D4ABC" };
            Rsid rsid42 = new Rsid() { Val = "007E0475" };
            Rsid rsid43 = new Rsid() { Val = "00822E25" };
            Rsid rsid44 = new Rsid() { Val = "00874D75" };
            Rsid rsid45 = new Rsid() { Val = "00895A72" };
            Rsid rsid46 = new Rsid() { Val = "008B4AAD" };
            Rsid rsid47 = new Rsid() { Val = "008D54AD" };
            Rsid rsid48 = new Rsid() { Val = "009056EA" };
            Rsid rsid49 = new Rsid() { Val = "009279AA" };
            Rsid rsid50 = new Rsid() { Val = "009758A4" };
            Rsid rsid51 = new Rsid() { Val = "009910CE" };
            Rsid rsid52 = new Rsid() { Val = "00995417" };
            Rsid rsid53 = new Rsid() { Val = "009E757C" };
            Rsid rsid54 = new Rsid() { Val = "009F15E5" };
            Rsid rsid55 = new Rsid() { Val = "009F467A" };
            Rsid rsid56 = new Rsid() { Val = "00A413B6" };
            Rsid rsid57 = new Rsid() { Val = "00B0242A" };
            Rsid rsid58 = new Rsid() { Val = "00B04547" };
            Rsid rsid59 = new Rsid() { Val = "00B15BB4" };
            Rsid rsid60 = new Rsid() { Val = "00B40558" };
            Rsid rsid61 = new Rsid() { Val = "00BA3CC1" };
            Rsid rsid62 = new Rsid() { Val = "00BB4A35" };
            Rsid rsid63 = new Rsid() { Val = "00BF4F85" };
            Rsid rsid64 = new Rsid() { Val = "00C22B76" };
            Rsid rsid65 = new Rsid() { Val = "00C454DF" };
            Rsid rsid66 = new Rsid() { Val = "00C747D0" };
            Rsid rsid67 = new Rsid() { Val = "00C8213A" };
            Rsid rsid68 = new Rsid() { Val = "00CB1952" };
            Rsid rsid69 = new Rsid() { Val = "00CD6E80" };
            Rsid rsid70 = new Rsid() { Val = "00D03B63" };
            Rsid rsid71 = new Rsid() { Val = "00D06B26" };
            Rsid rsid72 = new Rsid() { Val = "00D21E3F" };
            Rsid rsid73 = new Rsid() { Val = "00D43E79" };
            Rsid rsid74 = new Rsid() { Val = "00D730D5" };
            Rsid rsid75 = new Rsid() { Val = "00D8195F" };
            Rsid rsid76 = new Rsid() { Val = "00D9165B" };
            Rsid rsid77 = new Rsid() { Val = "00DA339C" };
            Rsid rsid78 = new Rsid() { Val = "00E103E6" };
            Rsid rsid79 = new Rsid() { Val = "00E15526" };
            Rsid rsid80 = new Rsid() { Val = "00E221A4" };
            Rsid rsid81 = new Rsid() { Val = "00E612F2" };
            Rsid rsid82 = new Rsid() { Val = "00E87996" };
            Rsid rsid83 = new Rsid() { Val = "00E95AFA" };
            Rsid rsid84 = new Rsid() { Val = "00ED6857" };
            Rsid rsid85 = new Rsid() { Val = "00EE2981" };
            Rsid rsid86 = new Rsid() { Val = "00EE72A8" };
            Rsid rsid87 = new Rsid() { Val = "00F73BAD" };
            Rsid rsid88 = new Rsid() { Val = "00F80A0B" };
            Rsid rsid89 = new Rsid() { Val = "00F8565B" };
            Rsid rsid90 = new Rsid() { Val = "00F901B4" };

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
            rsids1.Append(rsid46);
            rsids1.Append(rsid47);
            rsids1.Append(rsid48);
            rsids1.Append(rsid49);
            rsids1.Append(rsid50);
            rsids1.Append(rsid51);
            rsids1.Append(rsid52);
            rsids1.Append(rsid53);
            rsids1.Append(rsid54);
            rsids1.Append(rsid55);
            rsids1.Append(rsid56);
            rsids1.Append(rsid57);
            rsids1.Append(rsid58);
            rsids1.Append(rsid59);
            rsids1.Append(rsid60);
            rsids1.Append(rsid61);
            rsids1.Append(rsid62);
            rsids1.Append(rsid63);
            rsids1.Append(rsid64);
            rsids1.Append(rsid65);
            rsids1.Append(rsid66);
            rsids1.Append(rsid67);
            rsids1.Append(rsid68);
            rsids1.Append(rsid69);
            rsids1.Append(rsid70);
            rsids1.Append(rsid71);
            rsids1.Append(rsid72);
            rsids1.Append(rsid73);
            rsids1.Append(rsid74);
            rsids1.Append(rsid75);
            rsids1.Append(rsid76);
            rsids1.Append(rsid77);
            rsids1.Append(rsid78);
            rsids1.Append(rsid79);
            rsids1.Append(rsid80);
            rsids1.Append(rsid81);
            rsids1.Append(rsid82);
            rsids1.Append(rsid83);
            rsids1.Append(rsid84);
            rsids1.Append(rsid85);
            rsids1.Append(rsid86);
            rsids1.Append(rsid87);
            rsids1.Append(rsid88);
            rsids1.Append(rsid89);
            rsids1.Append(rsid90);

            M.MathProperties mathProperties1 = new M.MathProperties();
            M.MathFont mathFont1 = new M.MathFont() { Val = "Cambria Math" };
            M.BreakBinary breakBinary1 = new M.BreakBinary() { Val = M.BreakBinaryOperatorValues.Before };
            M.BreakBinarySubtraction breakBinarySubtraction1 = new M.BreakBinarySubtraction() { Val = M.BreakBinarySubtractionValues.MinusMinus };
            M.SmallFraction smallFraction1 = new M.SmallFraction() { Val = M.BooleanValues.Off };
            M.DisplayDefaults displayDefaults1 = new M.DisplayDefaults();
            M.LeftMargin leftMargin1 = new M.LeftMargin() { Val = (UInt32Value)0U };
            M.RightMargin rightMargin1 = new M.RightMargin() { Val = (UInt32Value)0U };
            M.DefaultJustification defaultJustification1 = new M.DefaultJustification() { Val = M.JustificationValues.CenterGroup };
            M.WrapIndent wrapIndent1 = new M.WrapIndent() { Val = (UInt32Value)1440U };
            M.IntegralLimitLocation integralLimitLocation1 = new M.IntegralLimitLocation() { Val = M.LimitLocationValues.SubscriptSuperscript };
            M.NaryLimitLocation naryLimitLocation1 = new M.NaryLimitLocation() { Val = M.LimitLocationValues.UnderOver };

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
            ThemeFontLanguages themeFontLanguages1 = new ThemeFontLanguages() { Val = "ru-RU" };
            ColorSchemeMapping colorSchemeMapping1 = new ColorSchemeMapping() { Background1 = ColorSchemeIndexValues.Light1, Text1 = ColorSchemeIndexValues.Dark1, Background2 = ColorSchemeIndexValues.Light2, Text2 = ColorSchemeIndexValues.Dark2, Accent1 = ColorSchemeIndexValues.Accent1, Accent2 = ColorSchemeIndexValues.Accent2, Accent3 = ColorSchemeIndexValues.Accent3, Accent4 = ColorSchemeIndexValues.Accent4, Accent5 = ColorSchemeIndexValues.Accent5, Accent6 = ColorSchemeIndexValues.Accent6, Hyperlink = ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = ColorSchemeIndexValues.FollowedHyperlink };
            DoNotIncludeSubdocsInStats doNotIncludeSubdocsInStats1 = new DoNotIncludeSubdocsInStats();
            DoNotAutoCompressPictures doNotAutoCompressPictures1 = new DoNotAutoCompressPictures();
            DecimalSymbol decimalSymbol1 = new DecimalSymbol() { Val = "," };
            ListSeparator listSeparator1 = new ListSeparator() { Val = ";" };

            settings1.Append(zoom1);
            settings1.Append(defaultTabStop1);
            settings1.Append(drawingGridHorizontalSpacing1);
            settings1.Append(drawingGridVerticalSpacing1);
            settings1.Append(characterSpacingControl1);
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

        // Generates content of styleDefinitionsPart1.
        private void GenerateStyleDefinitionsPart1Content(StyleDefinitionsPart styleDefinitionsPart1)
        {
            Styles styles1 = new Styles();
            styles1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            styles1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            DocDefaults docDefaults1 = new DocDefaults();

            RunPropertiesDefault runPropertiesDefault1 = new RunPropertiesDefault();

            RunPropertiesBaseStyle runPropertiesBaseStyle1 = new RunPropertiesBaseStyle();
            RunFonts runFonts3 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", EastAsia = "Calibri", ComplexScript = "Times New Roman" };
            FontSize fontSize3 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "22" };
            Languages languages1 = new Languages() { Val = "ru-RU", EastAsia = "ru-RU", Bidi = "ar-SA" };

            runPropertiesBaseStyle1.Append(runFonts3);
            runPropertiesBaseStyle1.Append(fontSize3);
            runPropertiesBaseStyle1.Append(fontSizeComplexScript3);
            runPropertiesBaseStyle1.Append(languages1);

            runPropertiesDefault1.Append(runPropertiesBaseStyle1);
            ParagraphPropertiesDefault paragraphPropertiesDefault1 = new ParagraphPropertiesDefault();

            docDefaults1.Append(runPropertiesDefault1);
            docDefaults1.Append(paragraphPropertiesDefault1);

            LatentStyles latentStyles1 = new LatentStyles() { DefaultLockedState = false, DefaultUiPriority = 99, DefaultSemiHidden = true, DefaultUnhideWhenUsed = true, DefaultPrimaryStyle = false, Count = 267 };
            LatentStyleExceptionInfo latentStyleExceptionInfo1 = new LatentStyleExceptionInfo() { Name = "Normal", Locked = true, UiPriority = 0, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo2 = new LatentStyleExceptionInfo() { Name = "heading 1", Locked = true, UiPriority = 0, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo3 = new LatentStyleExceptionInfo() { Name = "heading 2", Locked = true, UiPriority = 0, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo4 = new LatentStyleExceptionInfo() { Name = "heading 3", Locked = true, UiPriority = 0, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo5 = new LatentStyleExceptionInfo() { Name = "heading 4", Locked = true, UiPriority = 0, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo6 = new LatentStyleExceptionInfo() { Name = "heading 5", Locked = true, UiPriority = 0, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo7 = new LatentStyleExceptionInfo() { Name = "heading 6", Locked = true, UiPriority = 0, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo8 = new LatentStyleExceptionInfo() { Name = "heading 7", Locked = true, UiPriority = 0, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo9 = new LatentStyleExceptionInfo() { Name = "heading 8", Locked = true, UiPriority = 0, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo10 = new LatentStyleExceptionInfo() { Name = "heading 9", Locked = true, UiPriority = 0, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo11 = new LatentStyleExceptionInfo() { Name = "toc 1", Locked = true, UiPriority = 0, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo12 = new LatentStyleExceptionInfo() { Name = "toc 2", Locked = true, UiPriority = 0, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo13 = new LatentStyleExceptionInfo() { Name = "toc 3", Locked = true, UiPriority = 0, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo14 = new LatentStyleExceptionInfo() { Name = "toc 4", Locked = true, UiPriority = 0, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo15 = new LatentStyleExceptionInfo() { Name = "toc 5", Locked = true, UiPriority = 0, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo16 = new LatentStyleExceptionInfo() { Name = "toc 6", Locked = true, UiPriority = 0, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo17 = new LatentStyleExceptionInfo() { Name = "toc 7", Locked = true, UiPriority = 0, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo18 = new LatentStyleExceptionInfo() { Name = "toc 8", Locked = true, UiPriority = 0, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo19 = new LatentStyleExceptionInfo() { Name = "toc 9", Locked = true, UiPriority = 0, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo20 = new LatentStyleExceptionInfo() { Name = "caption", Locked = true, UiPriority = 0, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo21 = new LatentStyleExceptionInfo() { Name = "Title", Locked = true, UiPriority = 0, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo22 = new LatentStyleExceptionInfo() { Name = "Default Paragraph Font", Locked = true, UiPriority = 0, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo23 = new LatentStyleExceptionInfo() { Name = "Subtitle", Locked = true, UiPriority = 0, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo24 = new LatentStyleExceptionInfo() { Name = "Strong", Locked = true, UiPriority = 0, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo25 = new LatentStyleExceptionInfo() { Name = "Emphasis", Locked = true, UiPriority = 0, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo26 = new LatentStyleExceptionInfo() { Name = "Table Grid", Locked = true, UiPriority = 0, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo27 = new LatentStyleExceptionInfo() { Name = "Placeholder Text", UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo28 = new LatentStyleExceptionInfo() { Name = "No Spacing", UiPriority = 1, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo29 = new LatentStyleExceptionInfo() { Name = "Light Shading", UiPriority = 60, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo30 = new LatentStyleExceptionInfo() { Name = "Light List", UiPriority = 61, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo31 = new LatentStyleExceptionInfo() { Name = "Light Grid", UiPriority = 62, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo32 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1", UiPriority = 63, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo33 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2", UiPriority = 64, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo34 = new LatentStyleExceptionInfo() { Name = "Medium List 1", UiPriority = 65, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo35 = new LatentStyleExceptionInfo() { Name = "Medium List 2", UiPriority = 66, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo36 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1", UiPriority = 67, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo37 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2", UiPriority = 68, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo38 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3", UiPriority = 69, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo39 = new LatentStyleExceptionInfo() { Name = "Dark List", UiPriority = 70, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo40 = new LatentStyleExceptionInfo() { Name = "Colorful Shading", UiPriority = 71, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo41 = new LatentStyleExceptionInfo() { Name = "Colorful List", UiPriority = 72, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo42 = new LatentStyleExceptionInfo() { Name = "Colorful Grid", UiPriority = 73, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo43 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 1", UiPriority = 60, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo44 = new LatentStyleExceptionInfo() { Name = "Light List Accent 1", UiPriority = 61, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo45 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 1", UiPriority = 62, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo46 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 1", UiPriority = 63, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo47 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 1", UiPriority = 64, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo48 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 1", UiPriority = 65, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo49 = new LatentStyleExceptionInfo() { Name = "Revision", UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo50 = new LatentStyleExceptionInfo() { Name = "List Paragraph", UiPriority = 34, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo51 = new LatentStyleExceptionInfo() { Name = "Quote", UiPriority = 29, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo52 = new LatentStyleExceptionInfo() { Name = "Intense Quote", UiPriority = 30, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo53 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 1", UiPriority = 66, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo54 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 1", UiPriority = 67, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo55 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 1", UiPriority = 68, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo56 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 1", UiPriority = 69, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo57 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 1", UiPriority = 70, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo58 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 1", UiPriority = 71, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo59 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 1", UiPriority = 72, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo60 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 1", UiPriority = 73, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo61 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 2", UiPriority = 60, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo62 = new LatentStyleExceptionInfo() { Name = "Light List Accent 2", UiPriority = 61, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo63 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 2", UiPriority = 62, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo64 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 2", UiPriority = 63, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo65 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 2", UiPriority = 64, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo66 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 2", UiPriority = 65, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo67 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 2", UiPriority = 66, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo68 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 2", UiPriority = 67, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo69 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 2", UiPriority = 68, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo70 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 2", UiPriority = 69, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo71 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 2", UiPriority = 70, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo72 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 2", UiPriority = 71, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo73 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 2", UiPriority = 72, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo74 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 2", UiPriority = 73, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo75 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 3", UiPriority = 60, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo76 = new LatentStyleExceptionInfo() { Name = "Light List Accent 3", UiPriority = 61, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo77 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 3", UiPriority = 62, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo78 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 3", UiPriority = 63, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo79 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 3", UiPriority = 64, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo80 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 3", UiPriority = 65, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo81 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 3", UiPriority = 66, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo82 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 3", UiPriority = 67, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo83 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 3", UiPriority = 68, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo84 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 3", UiPriority = 69, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo85 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 3", UiPriority = 70, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo86 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 3", UiPriority = 71, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo87 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 3", UiPriority = 72, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo88 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 3", UiPriority = 73, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo89 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 4", UiPriority = 60, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo90 = new LatentStyleExceptionInfo() { Name = "Light List Accent 4", UiPriority = 61, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo91 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 4", UiPriority = 62, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo92 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 4", UiPriority = 63, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo93 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 4", UiPriority = 64, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo94 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 4", UiPriority = 65, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo95 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 4", UiPriority = 66, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo96 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 4", UiPriority = 67, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo97 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 4", UiPriority = 68, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo98 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 4", UiPriority = 69, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo99 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 4", UiPriority = 70, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo100 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 4", UiPriority = 71, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo101 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 4", UiPriority = 72, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo102 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 4", UiPriority = 73, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo103 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 5", UiPriority = 60, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo104 = new LatentStyleExceptionInfo() { Name = "Light List Accent 5", UiPriority = 61, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo105 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 5", UiPriority = 62, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo106 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 5", UiPriority = 63, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo107 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 5", UiPriority = 64, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo108 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 5", UiPriority = 65, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo109 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 5", UiPriority = 66, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo110 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 5", UiPriority = 67, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo111 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 5", UiPriority = 68, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo112 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 5", UiPriority = 69, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo113 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 5", UiPriority = 70, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo114 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 5", UiPriority = 71, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo115 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 5", UiPriority = 72, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo116 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 5", UiPriority = 73, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo117 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 6", UiPriority = 60, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo118 = new LatentStyleExceptionInfo() { Name = "Light List Accent 6", UiPriority = 61, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo119 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 6", UiPriority = 62, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo120 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 6", UiPriority = 63, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo121 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 6", UiPriority = 64, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo122 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 6", UiPriority = 65, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo123 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 6", UiPriority = 66, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo124 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 6", UiPriority = 67, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo125 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 6", UiPriority = 68, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo126 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 6", UiPriority = 69, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo127 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 6", UiPriority = 70, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo128 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 6", UiPriority = 71, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo129 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 6", UiPriority = 72, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo130 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 6", UiPriority = 73, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo131 = new LatentStyleExceptionInfo() { Name = "Subtle Emphasis", UiPriority = 19, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo132 = new LatentStyleExceptionInfo() { Name = "Intense Emphasis", UiPriority = 21, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo133 = new LatentStyleExceptionInfo() { Name = "Subtle Reference", UiPriority = 31, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo134 = new LatentStyleExceptionInfo() { Name = "Intense Reference", UiPriority = 32, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo135 = new LatentStyleExceptionInfo() { Name = "Book Title", UiPriority = 33, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo136 = new LatentStyleExceptionInfo() { Name = "Bibliography", UiPriority = 37 };
            LatentStyleExceptionInfo latentStyleExceptionInfo137 = new LatentStyleExceptionInfo() { Name = "TOC Heading", UiPriority = 39, PrimaryStyle = true };

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

            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true };
            StyleName styleName1 = new StyleName() { Val = "Normal" };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();
            Rsid rsid91 = new Rsid() { Val = "00D43E79" };

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines2 = new SpacingBetweenLines() { After = "160", Line = "256", LineRule = LineSpacingRuleValues.Auto };

            styleParagraphProperties1.Append(spacingBetweenLines2);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            Languages languages2 = new Languages() { EastAsia = "en-US" };

            styleRunProperties1.Append(languages2);

            style1.Append(styleName1);
            style1.Append(primaryStyle1);
            style1.Append(rsid91);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleRunProperties1);

            Style style2 = new Style() { Type = StyleValues.Character, StyleId = "DefaultParagraphFont", Default = true };
            StyleName styleName2 = new StyleName() { Val = "Default Paragraph Font" };
            UIPriority uIPriority1 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden1 = new SemiHidden();

            style2.Append(styleName2);
            style2.Append(uIPriority1);
            style2.Append(semiHidden1);

            Style style3 = new Style() { Type = StyleValues.Table, StyleId = "TableNormal", Default = true };
            StyleName styleName3 = new StyleName() { Val = "Normal Table" };
            UIPriority uIPriority2 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden2 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed1 = new UnhideWhenUsed();
            PrimaryStyle primaryStyle2 = new PrimaryStyle();

            StyleTableProperties styleTableProperties1 = new StyleTableProperties();
            TableIndentation tableIndentation1 = new TableIndentation() { Width = 0, Type = TableWidthUnitValues.Dxa };

            TableCellMarginDefault tableCellMarginDefault1 = new TableCellMarginDefault();
            TopMargin topMargin1 = new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellLeftMargin tableCellLeftMargin1 = new TableCellLeftMargin() { Width = 108, Type = TableWidthValues.Dxa };
            BottomMargin bottomMargin1 = new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellRightMargin tableCellRightMargin1 = new TableCellRightMargin() { Width = 108, Type = TableWidthValues.Dxa };

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

            Style style4 = new Style() { Type = StyleValues.Numbering, StyleId = "NoList", Default = true };
            StyleName styleName4 = new StyleName() { Val = "No List" };
            UIPriority uIPriority3 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden3 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed2 = new UnhideWhenUsed();

            style4.Append(styleName4);
            style4.Append(uIPriority3);
            style4.Append(semiHidden3);
            style4.Append(unhideWhenUsed2);

            Style style5 = new Style() { Type = StyleValues.Paragraph, StyleId = "BalloonText" };
            StyleName styleName5 = new StyleName() { Val = "Balloon Text" };
            BasedOn basedOn1 = new BasedOn() { Val = "Normal" };
            LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "BalloonTextChar" };
            UIPriority uIPriority4 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden4 = new SemiHidden();
            Rsid rsid92 = new Rsid() { Val = "00822E25" };

            StyleParagraphProperties styleParagraphProperties2 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines3 = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };

            styleParagraphProperties2.Append(spacingBetweenLines3);

            StyleRunProperties styleRunProperties2 = new StyleRunProperties();
            RunFonts runFonts4 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            FontSize fontSize4 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "16" };

            styleRunProperties2.Append(runFonts4);
            styleRunProperties2.Append(fontSize4);
            styleRunProperties2.Append(fontSizeComplexScript4);

            style5.Append(styleName5);
            style5.Append(basedOn1);
            style5.Append(linkedStyle1);
            style5.Append(uIPriority4);
            style5.Append(semiHidden4);
            style5.Append(rsid92);
            style5.Append(styleParagraphProperties2);
            style5.Append(styleRunProperties2);

            Style style6 = new Style() { Type = StyleValues.Character, StyleId = "BalloonTextChar", CustomStyle = true };
            StyleName styleName6 = new StyleName() { Val = "Balloon Text Char" };
            BasedOn basedOn2 = new BasedOn() { Val = "DefaultParagraphFont" };
            LinkedStyle linkedStyle2 = new LinkedStyle() { Val = "BalloonText" };
            UIPriority uIPriority5 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden5 = new SemiHidden();
            Locked locked1 = new Locked();
            Rsid rsid93 = new Rsid() { Val = "00822E25" };

            StyleRunProperties styleRunProperties3 = new StyleRunProperties();
            RunFonts runFonts5 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            FontSize fontSize5 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript5 = new FontSizeComplexScript() { Val = "16" };

            styleRunProperties3.Append(runFonts5);
            styleRunProperties3.Append(fontSize5);
            styleRunProperties3.Append(fontSizeComplexScript5);

            style6.Append(styleName6);
            style6.Append(basedOn2);
            style6.Append(linkedStyle2);
            style6.Append(uIPriority5);
            style6.Append(semiHidden5);
            style6.Append(locked1);
            style6.Append(rsid93);
            style6.Append(styleRunProperties3);

            Style style7 = new Style() { Type = StyleValues.Paragraph, StyleId = "ListParagraph" };
            StyleName styleName7 = new StyleName() { Val = "List Paragraph" };
            BasedOn basedOn3 = new BasedOn() { Val = "Normal" };
            UIPriority uIPriority6 = new UIPriority() { Val = 99 };
            PrimaryStyle primaryStyle3 = new PrimaryStyle();
            Rsid rsid94 = new Rsid() { Val = "00504C5D" };

            StyleParagraphProperties styleParagraphProperties3 = new StyleParagraphProperties();
            Indentation indentation1 = new Indentation() { Left = "720" };
            ContextualSpacing contextualSpacing1 = new ContextualSpacing();

            styleParagraphProperties3.Append(indentation1);
            styleParagraphProperties3.Append(contextualSpacing1);

            style7.Append(styleName7);
            style7.Append(basedOn3);
            style7.Append(uIPriority6);
            style7.Append(primaryStyle3);
            style7.Append(rsid94);
            style7.Append(styleParagraphProperties3);

            Style style8 = new Style() { Type = StyleValues.Character, StyleId = "PlaceholderText" };
            StyleName styleName8 = new StyleName() { Val = "Placeholder Text" };
            BasedOn basedOn4 = new BasedOn() { Val = "DefaultParagraphFont" };
            UIPriority uIPriority7 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden6 = new SemiHidden();
            Rsid rsid95 = new Rsid() { Val = "002B2FE4" };

            StyleRunProperties styleRunProperties4 = new StyleRunProperties();
            RunFonts runFonts6 = new RunFonts() { ComplexScript = "Times New Roman" };
            Color color1 = new Color() { Val = "808080" };

            styleRunProperties4.Append(runFonts6);
            styleRunProperties4.Append(color1);

            style8.Append(styleName8);
            style8.Append(basedOn4);
            style8.Append(uIPriority7);
            style8.Append(semiHidden6);
            style8.Append(rsid95);
            style8.Append(styleRunProperties4);

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

            AbstractNum abstractNum1 = new AbstractNum() { AbstractNumberId = 0 };
            Nsid nsid1 = new Nsid() { Val = "22711487" };
            MultiLevelType multiLevelType1 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
            TemplateCode templateCode1 = new TemplateCode() { Val = "D50EF818" };

            Level level1 = new Level() { LevelIndex = 0, TemplateCode = "04190011" };
            StartNumberingValue startNumberingValue1 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat1 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText1 = new LevelText() { Val = "%1)" };
            LevelJustification levelJustification1 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties1 = new PreviousParagraphProperties();
            Indentation indentation2 = new Indentation() { Left = "720", Hanging = "360" };

            previousParagraphProperties1.Append(indentation2);

            NumberingSymbolRunProperties numberingSymbolRunProperties1 = new NumberingSymbolRunProperties();
            RunFonts runFonts7 = new RunFonts() { Hint = FontTypeHintValues.Default, ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties1.Append(runFonts7);

            level1.Append(startNumberingValue1);
            level1.Append(numberingFormat1);
            level1.Append(levelText1);
            level1.Append(levelJustification1);
            level1.Append(previousParagraphProperties1);
            level1.Append(numberingSymbolRunProperties1);

            Level level2 = new Level() { LevelIndex = 1, TemplateCode = "04190019", Tentative = true };
            StartNumberingValue startNumberingValue2 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat2 = new NumberingFormat() { Val = NumberFormatValues.LowerLetter };
            LevelText levelText2 = new LevelText() { Val = "%2." };
            LevelJustification levelJustification2 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties2 = new PreviousParagraphProperties();
            Indentation indentation3 = new Indentation() { Left = "1440", Hanging = "360" };

            previousParagraphProperties2.Append(indentation3);

            NumberingSymbolRunProperties numberingSymbolRunProperties2 = new NumberingSymbolRunProperties();
            RunFonts runFonts8 = new RunFonts() { ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties2.Append(runFonts8);

            level2.Append(startNumberingValue2);
            level2.Append(numberingFormat2);
            level2.Append(levelText2);
            level2.Append(levelJustification2);
            level2.Append(previousParagraphProperties2);
            level2.Append(numberingSymbolRunProperties2);

            Level level3 = new Level() { LevelIndex = 2, TemplateCode = "0419001B", Tentative = true };
            StartNumberingValue startNumberingValue3 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat3 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText3 = new LevelText() { Val = "%3." };
            LevelJustification levelJustification3 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties3 = new PreviousParagraphProperties();
            Indentation indentation4 = new Indentation() { Left = "2160", Hanging = "180" };

            previousParagraphProperties3.Append(indentation4);

            NumberingSymbolRunProperties numberingSymbolRunProperties3 = new NumberingSymbolRunProperties();
            RunFonts runFonts9 = new RunFonts() { ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties3.Append(runFonts9);

            level3.Append(startNumberingValue3);
            level3.Append(numberingFormat3);
            level3.Append(levelText3);
            level3.Append(levelJustification3);
            level3.Append(previousParagraphProperties3);
            level3.Append(numberingSymbolRunProperties3);

            Level level4 = new Level() { LevelIndex = 3, TemplateCode = "0419000F", Tentative = true };
            StartNumberingValue startNumberingValue4 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat4 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText4 = new LevelText() { Val = "%4." };
            LevelJustification levelJustification4 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties4 = new PreviousParagraphProperties();
            Indentation indentation5 = new Indentation() { Left = "2880", Hanging = "360" };

            previousParagraphProperties4.Append(indentation5);

            NumberingSymbolRunProperties numberingSymbolRunProperties4 = new NumberingSymbolRunProperties();
            RunFonts runFonts10 = new RunFonts() { ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties4.Append(runFonts10);

            level4.Append(startNumberingValue4);
            level4.Append(numberingFormat4);
            level4.Append(levelText4);
            level4.Append(levelJustification4);
            level4.Append(previousParagraphProperties4);
            level4.Append(numberingSymbolRunProperties4);

            Level level5 = new Level() { LevelIndex = 4, TemplateCode = "04190019", Tentative = true };
            StartNumberingValue startNumberingValue5 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat5 = new NumberingFormat() { Val = NumberFormatValues.LowerLetter };
            LevelText levelText5 = new LevelText() { Val = "%5." };
            LevelJustification levelJustification5 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties5 = new PreviousParagraphProperties();
            Indentation indentation6 = new Indentation() { Left = "3600", Hanging = "360" };

            previousParagraphProperties5.Append(indentation6);

            NumberingSymbolRunProperties numberingSymbolRunProperties5 = new NumberingSymbolRunProperties();
            RunFonts runFonts11 = new RunFonts() { ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties5.Append(runFonts11);

            level5.Append(startNumberingValue5);
            level5.Append(numberingFormat5);
            level5.Append(levelText5);
            level5.Append(levelJustification5);
            level5.Append(previousParagraphProperties5);
            level5.Append(numberingSymbolRunProperties5);

            Level level6 = new Level() { LevelIndex = 5, TemplateCode = "0419001B", Tentative = true };
            StartNumberingValue startNumberingValue6 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat6 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText6 = new LevelText() { Val = "%6." };
            LevelJustification levelJustification6 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties6 = new PreviousParagraphProperties();
            Indentation indentation7 = new Indentation() { Left = "4320", Hanging = "180" };

            previousParagraphProperties6.Append(indentation7);

            NumberingSymbolRunProperties numberingSymbolRunProperties6 = new NumberingSymbolRunProperties();
            RunFonts runFonts12 = new RunFonts() { ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties6.Append(runFonts12);

            level6.Append(startNumberingValue6);
            level6.Append(numberingFormat6);
            level6.Append(levelText6);
            level6.Append(levelJustification6);
            level6.Append(previousParagraphProperties6);
            level6.Append(numberingSymbolRunProperties6);

            Level level7 = new Level() { LevelIndex = 6, TemplateCode = "0419000F", Tentative = true };
            StartNumberingValue startNumberingValue7 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat7 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText7 = new LevelText() { Val = "%7." };
            LevelJustification levelJustification7 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties7 = new PreviousParagraphProperties();
            Indentation indentation8 = new Indentation() { Left = "5040", Hanging = "360" };

            previousParagraphProperties7.Append(indentation8);

            NumberingSymbolRunProperties numberingSymbolRunProperties7 = new NumberingSymbolRunProperties();
            RunFonts runFonts13 = new RunFonts() { ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties7.Append(runFonts13);

            level7.Append(startNumberingValue7);
            level7.Append(numberingFormat7);
            level7.Append(levelText7);
            level7.Append(levelJustification7);
            level7.Append(previousParagraphProperties7);
            level7.Append(numberingSymbolRunProperties7);

            Level level8 = new Level() { LevelIndex = 7, TemplateCode = "04190019", Tentative = true };
            StartNumberingValue startNumberingValue8 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat8 = new NumberingFormat() { Val = NumberFormatValues.LowerLetter };
            LevelText levelText8 = new LevelText() { Val = "%8." };
            LevelJustification levelJustification8 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties8 = new PreviousParagraphProperties();
            Indentation indentation9 = new Indentation() { Left = "5760", Hanging = "360" };

            previousParagraphProperties8.Append(indentation9);

            NumberingSymbolRunProperties numberingSymbolRunProperties8 = new NumberingSymbolRunProperties();
            RunFonts runFonts14 = new RunFonts() { ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties8.Append(runFonts14);

            level8.Append(startNumberingValue8);
            level8.Append(numberingFormat8);
            level8.Append(levelText8);
            level8.Append(levelJustification8);
            level8.Append(previousParagraphProperties8);
            level8.Append(numberingSymbolRunProperties8);

            Level level9 = new Level() { LevelIndex = 8, TemplateCode = "0419001B", Tentative = true };
            StartNumberingValue startNumberingValue9 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat9 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText9 = new LevelText() { Val = "%9." };
            LevelJustification levelJustification9 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties9 = new PreviousParagraphProperties();
            Indentation indentation10 = new Indentation() { Left = "6480", Hanging = "180" };

            previousParagraphProperties9.Append(indentation10);

            NumberingSymbolRunProperties numberingSymbolRunProperties9 = new NumberingSymbolRunProperties();
            RunFonts runFonts15 = new RunFonts() { ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties9.Append(runFonts15);

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

            AbstractNum abstractNum2 = new AbstractNum() { AbstractNumberId = 1 };
            Nsid nsid2 = new Nsid() { Val = "2C47498D" };
            MultiLevelType multiLevelType2 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
            TemplateCode templateCode2 = new TemplateCode() { Val = "6524A47A" };

            Level level10 = new Level() { LevelIndex = 0, TemplateCode = "04190011" };
            StartNumberingValue startNumberingValue10 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat10 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText10 = new LevelText() { Val = "%1)" };
            LevelJustification levelJustification10 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties10 = new PreviousParagraphProperties();
            Indentation indentation11 = new Indentation() { Left = "720", Hanging = "360" };

            previousParagraphProperties10.Append(indentation11);

            NumberingSymbolRunProperties numberingSymbolRunProperties10 = new NumberingSymbolRunProperties();
            RunFonts runFonts16 = new RunFonts() { Hint = FontTypeHintValues.Default, ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties10.Append(runFonts16);

            level10.Append(startNumberingValue10);
            level10.Append(numberingFormat10);
            level10.Append(levelText10);
            level10.Append(levelJustification10);
            level10.Append(previousParagraphProperties10);
            level10.Append(numberingSymbolRunProperties10);

            Level level11 = new Level() { LevelIndex = 1, TemplateCode = "04190019", Tentative = true };
            StartNumberingValue startNumberingValue11 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat11 = new NumberingFormat() { Val = NumberFormatValues.LowerLetter };
            LevelText levelText11 = new LevelText() { Val = "%2." };
            LevelJustification levelJustification11 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties11 = new PreviousParagraphProperties();
            Indentation indentation12 = new Indentation() { Left = "1440", Hanging = "360" };

            previousParagraphProperties11.Append(indentation12);

            NumberingSymbolRunProperties numberingSymbolRunProperties11 = new NumberingSymbolRunProperties();
            RunFonts runFonts17 = new RunFonts() { ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties11.Append(runFonts17);

            level11.Append(startNumberingValue11);
            level11.Append(numberingFormat11);
            level11.Append(levelText11);
            level11.Append(levelJustification11);
            level11.Append(previousParagraphProperties11);
            level11.Append(numberingSymbolRunProperties11);

            Level level12 = new Level() { LevelIndex = 2, TemplateCode = "0419001B", Tentative = true };
            StartNumberingValue startNumberingValue12 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat12 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText12 = new LevelText() { Val = "%3." };
            LevelJustification levelJustification12 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties12 = new PreviousParagraphProperties();
            Indentation indentation13 = new Indentation() { Left = "2160", Hanging = "180" };

            previousParagraphProperties12.Append(indentation13);

            NumberingSymbolRunProperties numberingSymbolRunProperties12 = new NumberingSymbolRunProperties();
            RunFonts runFonts18 = new RunFonts() { ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties12.Append(runFonts18);

            level12.Append(startNumberingValue12);
            level12.Append(numberingFormat12);
            level12.Append(levelText12);
            level12.Append(levelJustification12);
            level12.Append(previousParagraphProperties12);
            level12.Append(numberingSymbolRunProperties12);

            Level level13 = new Level() { LevelIndex = 3, TemplateCode = "0419000F", Tentative = true };
            StartNumberingValue startNumberingValue13 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat13 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText13 = new LevelText() { Val = "%4." };
            LevelJustification levelJustification13 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties13 = new PreviousParagraphProperties();
            Indentation indentation14 = new Indentation() { Left = "2880", Hanging = "360" };

            previousParagraphProperties13.Append(indentation14);

            NumberingSymbolRunProperties numberingSymbolRunProperties13 = new NumberingSymbolRunProperties();
            RunFonts runFonts19 = new RunFonts() { ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties13.Append(runFonts19);

            level13.Append(startNumberingValue13);
            level13.Append(numberingFormat13);
            level13.Append(levelText13);
            level13.Append(levelJustification13);
            level13.Append(previousParagraphProperties13);
            level13.Append(numberingSymbolRunProperties13);

            Level level14 = new Level() { LevelIndex = 4, TemplateCode = "04190019", Tentative = true };
            StartNumberingValue startNumberingValue14 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat14 = new NumberingFormat() { Val = NumberFormatValues.LowerLetter };
            LevelText levelText14 = new LevelText() { Val = "%5." };
            LevelJustification levelJustification14 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties14 = new PreviousParagraphProperties();
            Indentation indentation15 = new Indentation() { Left = "3600", Hanging = "360" };

            previousParagraphProperties14.Append(indentation15);

            NumberingSymbolRunProperties numberingSymbolRunProperties14 = new NumberingSymbolRunProperties();
            RunFonts runFonts20 = new RunFonts() { ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties14.Append(runFonts20);

            level14.Append(startNumberingValue14);
            level14.Append(numberingFormat14);
            level14.Append(levelText14);
            level14.Append(levelJustification14);
            level14.Append(previousParagraphProperties14);
            level14.Append(numberingSymbolRunProperties14);

            Level level15 = new Level() { LevelIndex = 5, TemplateCode = "0419001B", Tentative = true };
            StartNumberingValue startNumberingValue15 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat15 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText15 = new LevelText() { Val = "%6." };
            LevelJustification levelJustification15 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties15 = new PreviousParagraphProperties();
            Indentation indentation16 = new Indentation() { Left = "4320", Hanging = "180" };

            previousParagraphProperties15.Append(indentation16);

            NumberingSymbolRunProperties numberingSymbolRunProperties15 = new NumberingSymbolRunProperties();
            RunFonts runFonts21 = new RunFonts() { ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties15.Append(runFonts21);

            level15.Append(startNumberingValue15);
            level15.Append(numberingFormat15);
            level15.Append(levelText15);
            level15.Append(levelJustification15);
            level15.Append(previousParagraphProperties15);
            level15.Append(numberingSymbolRunProperties15);

            Level level16 = new Level() { LevelIndex = 6, TemplateCode = "0419000F", Tentative = true };
            StartNumberingValue startNumberingValue16 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat16 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText16 = new LevelText() { Val = "%7." };
            LevelJustification levelJustification16 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties16 = new PreviousParagraphProperties();
            Indentation indentation17 = new Indentation() { Left = "5040", Hanging = "360" };

            previousParagraphProperties16.Append(indentation17);

            NumberingSymbolRunProperties numberingSymbolRunProperties16 = new NumberingSymbolRunProperties();
            RunFonts runFonts22 = new RunFonts() { ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties16.Append(runFonts22);

            level16.Append(startNumberingValue16);
            level16.Append(numberingFormat16);
            level16.Append(levelText16);
            level16.Append(levelJustification16);
            level16.Append(previousParagraphProperties16);
            level16.Append(numberingSymbolRunProperties16);

            Level level17 = new Level() { LevelIndex = 7, TemplateCode = "04190019", Tentative = true };
            StartNumberingValue startNumberingValue17 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat17 = new NumberingFormat() { Val = NumberFormatValues.LowerLetter };
            LevelText levelText17 = new LevelText() { Val = "%8." };
            LevelJustification levelJustification17 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties17 = new PreviousParagraphProperties();
            Indentation indentation18 = new Indentation() { Left = "5760", Hanging = "360" };

            previousParagraphProperties17.Append(indentation18);

            NumberingSymbolRunProperties numberingSymbolRunProperties17 = new NumberingSymbolRunProperties();
            RunFonts runFonts23 = new RunFonts() { ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties17.Append(runFonts23);

            level17.Append(startNumberingValue17);
            level17.Append(numberingFormat17);
            level17.Append(levelText17);
            level17.Append(levelJustification17);
            level17.Append(previousParagraphProperties17);
            level17.Append(numberingSymbolRunProperties17);

            Level level18 = new Level() { LevelIndex = 8, TemplateCode = "0419001B", Tentative = true };
            StartNumberingValue startNumberingValue18 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat18 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText18 = new LevelText() { Val = "%9." };
            LevelJustification levelJustification18 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties18 = new PreviousParagraphProperties();
            Indentation indentation19 = new Indentation() { Left = "6480", Hanging = "180" };

            previousParagraphProperties18.Append(indentation19);

            NumberingSymbolRunProperties numberingSymbolRunProperties18 = new NumberingSymbolRunProperties();
            RunFonts runFonts24 = new RunFonts() { ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties18.Append(runFonts24);

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

            AbstractNum abstractNum3 = new AbstractNum() { AbstractNumberId = 2 };
            Nsid nsid3 = new Nsid() { Val = "454B11A6" };
            MultiLevelType multiLevelType3 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
            TemplateCode templateCode3 = new TemplateCode() { Val = "30463A3A" };

            Level level19 = new Level() { LevelIndex = 0, TemplateCode = "04190011" };
            StartNumberingValue startNumberingValue19 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat19 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText19 = new LevelText() { Val = "%1)" };
            LevelJustification levelJustification19 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties19 = new PreviousParagraphProperties();
            Indentation indentation20 = new Indentation() { Left = "720", Hanging = "360" };

            previousParagraphProperties19.Append(indentation20);

            NumberingSymbolRunProperties numberingSymbolRunProperties19 = new NumberingSymbolRunProperties();
            RunFonts runFonts25 = new RunFonts() { Hint = FontTypeHintValues.Default, ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties19.Append(runFonts25);

            level19.Append(startNumberingValue19);
            level19.Append(numberingFormat19);
            level19.Append(levelText19);
            level19.Append(levelJustification19);
            level19.Append(previousParagraphProperties19);
            level19.Append(numberingSymbolRunProperties19);

            Level level20 = new Level() { LevelIndex = 1, TemplateCode = "04190019", Tentative = true };
            StartNumberingValue startNumberingValue20 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat20 = new NumberingFormat() { Val = NumberFormatValues.LowerLetter };
            LevelText levelText20 = new LevelText() { Val = "%2." };
            LevelJustification levelJustification20 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties20 = new PreviousParagraphProperties();
            Indentation indentation21 = new Indentation() { Left = "1440", Hanging = "360" };

            previousParagraphProperties20.Append(indentation21);

            NumberingSymbolRunProperties numberingSymbolRunProperties20 = new NumberingSymbolRunProperties();
            RunFonts runFonts26 = new RunFonts() { ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties20.Append(runFonts26);

            level20.Append(startNumberingValue20);
            level20.Append(numberingFormat20);
            level20.Append(levelText20);
            level20.Append(levelJustification20);
            level20.Append(previousParagraphProperties20);
            level20.Append(numberingSymbolRunProperties20);

            Level level21 = new Level() { LevelIndex = 2, TemplateCode = "0419001B", Tentative = true };
            StartNumberingValue startNumberingValue21 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat21 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText21 = new LevelText() { Val = "%3." };
            LevelJustification levelJustification21 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties21 = new PreviousParagraphProperties();
            Indentation indentation22 = new Indentation() { Left = "2160", Hanging = "180" };

            previousParagraphProperties21.Append(indentation22);

            NumberingSymbolRunProperties numberingSymbolRunProperties21 = new NumberingSymbolRunProperties();
            RunFonts runFonts27 = new RunFonts() { ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties21.Append(runFonts27);

            level21.Append(startNumberingValue21);
            level21.Append(numberingFormat21);
            level21.Append(levelText21);
            level21.Append(levelJustification21);
            level21.Append(previousParagraphProperties21);
            level21.Append(numberingSymbolRunProperties21);

            Level level22 = new Level() { LevelIndex = 3, TemplateCode = "0419000F", Tentative = true };
            StartNumberingValue startNumberingValue22 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat22 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText22 = new LevelText() { Val = "%4." };
            LevelJustification levelJustification22 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties22 = new PreviousParagraphProperties();
            Indentation indentation23 = new Indentation() { Left = "2880", Hanging = "360" };

            previousParagraphProperties22.Append(indentation23);

            NumberingSymbolRunProperties numberingSymbolRunProperties22 = new NumberingSymbolRunProperties();
            RunFonts runFonts28 = new RunFonts() { ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties22.Append(runFonts28);

            level22.Append(startNumberingValue22);
            level22.Append(numberingFormat22);
            level22.Append(levelText22);
            level22.Append(levelJustification22);
            level22.Append(previousParagraphProperties22);
            level22.Append(numberingSymbolRunProperties22);

            Level level23 = new Level() { LevelIndex = 4, TemplateCode = "04190019", Tentative = true };
            StartNumberingValue startNumberingValue23 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat23 = new NumberingFormat() { Val = NumberFormatValues.LowerLetter };
            LevelText levelText23 = new LevelText() { Val = "%5." };
            LevelJustification levelJustification23 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties23 = new PreviousParagraphProperties();
            Indentation indentation24 = new Indentation() { Left = "3600", Hanging = "360" };

            previousParagraphProperties23.Append(indentation24);

            NumberingSymbolRunProperties numberingSymbolRunProperties23 = new NumberingSymbolRunProperties();
            RunFonts runFonts29 = new RunFonts() { ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties23.Append(runFonts29);

            level23.Append(startNumberingValue23);
            level23.Append(numberingFormat23);
            level23.Append(levelText23);
            level23.Append(levelJustification23);
            level23.Append(previousParagraphProperties23);
            level23.Append(numberingSymbolRunProperties23);

            Level level24 = new Level() { LevelIndex = 5, TemplateCode = "0419001B", Tentative = true };
            StartNumberingValue startNumberingValue24 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat24 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText24 = new LevelText() { Val = "%6." };
            LevelJustification levelJustification24 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties24 = new PreviousParagraphProperties();
            Indentation indentation25 = new Indentation() { Left = "4320", Hanging = "180" };

            previousParagraphProperties24.Append(indentation25);

            NumberingSymbolRunProperties numberingSymbolRunProperties24 = new NumberingSymbolRunProperties();
            RunFonts runFonts30 = new RunFonts() { ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties24.Append(runFonts30);

            level24.Append(startNumberingValue24);
            level24.Append(numberingFormat24);
            level24.Append(levelText24);
            level24.Append(levelJustification24);
            level24.Append(previousParagraphProperties24);
            level24.Append(numberingSymbolRunProperties24);

            Level level25 = new Level() { LevelIndex = 6, TemplateCode = "0419000F", Tentative = true };
            StartNumberingValue startNumberingValue25 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat25 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText25 = new LevelText() { Val = "%7." };
            LevelJustification levelJustification25 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties25 = new PreviousParagraphProperties();
            Indentation indentation26 = new Indentation() { Left = "5040", Hanging = "360" };

            previousParagraphProperties25.Append(indentation26);

            NumberingSymbolRunProperties numberingSymbolRunProperties25 = new NumberingSymbolRunProperties();
            RunFonts runFonts31 = new RunFonts() { ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties25.Append(runFonts31);

            level25.Append(startNumberingValue25);
            level25.Append(numberingFormat25);
            level25.Append(levelText25);
            level25.Append(levelJustification25);
            level25.Append(previousParagraphProperties25);
            level25.Append(numberingSymbolRunProperties25);

            Level level26 = new Level() { LevelIndex = 7, TemplateCode = "04190019", Tentative = true };
            StartNumberingValue startNumberingValue26 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat26 = new NumberingFormat() { Val = NumberFormatValues.LowerLetter };
            LevelText levelText26 = new LevelText() { Val = "%8." };
            LevelJustification levelJustification26 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties26 = new PreviousParagraphProperties();
            Indentation indentation27 = new Indentation() { Left = "5760", Hanging = "360" };

            previousParagraphProperties26.Append(indentation27);

            NumberingSymbolRunProperties numberingSymbolRunProperties26 = new NumberingSymbolRunProperties();
            RunFonts runFonts32 = new RunFonts() { ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties26.Append(runFonts32);

            level26.Append(startNumberingValue26);
            level26.Append(numberingFormat26);
            level26.Append(levelText26);
            level26.Append(levelJustification26);
            level26.Append(previousParagraphProperties26);
            level26.Append(numberingSymbolRunProperties26);

            Level level27 = new Level() { LevelIndex = 8, TemplateCode = "0419001B", Tentative = true };
            StartNumberingValue startNumberingValue27 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat27 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText27 = new LevelText() { Val = "%9." };
            LevelJustification levelJustification27 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties27 = new PreviousParagraphProperties();
            Indentation indentation28 = new Indentation() { Left = "6480", Hanging = "180" };

            previousParagraphProperties27.Append(indentation28);

            NumberingSymbolRunProperties numberingSymbolRunProperties27 = new NumberingSymbolRunProperties();
            RunFonts runFonts33 = new RunFonts() { ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties27.Append(runFonts33);

            level27.Append(startNumberingValue27);
            level27.Append(numberingFormat27);
            level27.Append(levelText27);
            level27.Append(levelJustification27);
            level27.Append(previousParagraphProperties27);
            level27.Append(numberingSymbolRunProperties27);

            abstractNum3.Append(nsid3);
            abstractNum3.Append(multiLevelType3);
            abstractNum3.Append(templateCode3);
            abstractNum3.Append(level19);
            abstractNum3.Append(level20);
            abstractNum3.Append(level21);
            abstractNum3.Append(level22);
            abstractNum3.Append(level23);
            abstractNum3.Append(level24);
            abstractNum3.Append(level25);
            abstractNum3.Append(level26);
            abstractNum3.Append(level27);

            AbstractNum abstractNum4 = new AbstractNum() { AbstractNumberId = 3 };
            Nsid nsid4 = new Nsid() { Val = "5E2200FA" };
            MultiLevelType multiLevelType4 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
            TemplateCode templateCode4 = new TemplateCode() { Val = "6444E6EC" };

            Level level28 = new Level() { LevelIndex = 0, TemplateCode = "04190011" };
            StartNumberingValue startNumberingValue28 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat28 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText28 = new LevelText() { Val = "%1)" };
            LevelJustification levelJustification28 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties28 = new PreviousParagraphProperties();
            Indentation indentation29 = new Indentation() { Left = "720", Hanging = "360" };

            previousParagraphProperties28.Append(indentation29);

            NumberingSymbolRunProperties numberingSymbolRunProperties28 = new NumberingSymbolRunProperties();
            RunFonts runFonts34 = new RunFonts() { Hint = FontTypeHintValues.Default, ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties28.Append(runFonts34);

            level28.Append(startNumberingValue28);
            level28.Append(numberingFormat28);
            level28.Append(levelText28);
            level28.Append(levelJustification28);
            level28.Append(previousParagraphProperties28);
            level28.Append(numberingSymbolRunProperties28);

            Level level29 = new Level() { LevelIndex = 1, TemplateCode = "04190019", Tentative = true };
            StartNumberingValue startNumberingValue29 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat29 = new NumberingFormat() { Val = NumberFormatValues.LowerLetter };
            LevelText levelText29 = new LevelText() { Val = "%2." };
            LevelJustification levelJustification29 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties29 = new PreviousParagraphProperties();
            Indentation indentation30 = new Indentation() { Left = "1440", Hanging = "360" };

            previousParagraphProperties29.Append(indentation30);

            NumberingSymbolRunProperties numberingSymbolRunProperties29 = new NumberingSymbolRunProperties();
            RunFonts runFonts35 = new RunFonts() { ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties29.Append(runFonts35);

            level29.Append(startNumberingValue29);
            level29.Append(numberingFormat29);
            level29.Append(levelText29);
            level29.Append(levelJustification29);
            level29.Append(previousParagraphProperties29);
            level29.Append(numberingSymbolRunProperties29);

            Level level30 = new Level() { LevelIndex = 2, TemplateCode = "0419001B", Tentative = true };
            StartNumberingValue startNumberingValue30 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat30 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText30 = new LevelText() { Val = "%3." };
            LevelJustification levelJustification30 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties30 = new PreviousParagraphProperties();
            Indentation indentation31 = new Indentation() { Left = "2160", Hanging = "180" };

            previousParagraphProperties30.Append(indentation31);

            NumberingSymbolRunProperties numberingSymbolRunProperties30 = new NumberingSymbolRunProperties();
            RunFonts runFonts36 = new RunFonts() { ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties30.Append(runFonts36);

            level30.Append(startNumberingValue30);
            level30.Append(numberingFormat30);
            level30.Append(levelText30);
            level30.Append(levelJustification30);
            level30.Append(previousParagraphProperties30);
            level30.Append(numberingSymbolRunProperties30);

            Level level31 = new Level() { LevelIndex = 3, TemplateCode = "0419000F", Tentative = true };
            StartNumberingValue startNumberingValue31 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat31 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText31 = new LevelText() { Val = "%4." };
            LevelJustification levelJustification31 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties31 = new PreviousParagraphProperties();
            Indentation indentation32 = new Indentation() { Left = "2880", Hanging = "360" };

            previousParagraphProperties31.Append(indentation32);

            NumberingSymbolRunProperties numberingSymbolRunProperties31 = new NumberingSymbolRunProperties();
            RunFonts runFonts37 = new RunFonts() { ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties31.Append(runFonts37);

            level31.Append(startNumberingValue31);
            level31.Append(numberingFormat31);
            level31.Append(levelText31);
            level31.Append(levelJustification31);
            level31.Append(previousParagraphProperties31);
            level31.Append(numberingSymbolRunProperties31);

            Level level32 = new Level() { LevelIndex = 4, TemplateCode = "04190019", Tentative = true };
            StartNumberingValue startNumberingValue32 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat32 = new NumberingFormat() { Val = NumberFormatValues.LowerLetter };
            LevelText levelText32 = new LevelText() { Val = "%5." };
            LevelJustification levelJustification32 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties32 = new PreviousParagraphProperties();
            Indentation indentation33 = new Indentation() { Left = "3600", Hanging = "360" };

            previousParagraphProperties32.Append(indentation33);

            NumberingSymbolRunProperties numberingSymbolRunProperties32 = new NumberingSymbolRunProperties();
            RunFonts runFonts38 = new RunFonts() { ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties32.Append(runFonts38);

            level32.Append(startNumberingValue32);
            level32.Append(numberingFormat32);
            level32.Append(levelText32);
            level32.Append(levelJustification32);
            level32.Append(previousParagraphProperties32);
            level32.Append(numberingSymbolRunProperties32);

            Level level33 = new Level() { LevelIndex = 5, TemplateCode = "0419001B", Tentative = true };
            StartNumberingValue startNumberingValue33 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat33 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText33 = new LevelText() { Val = "%6." };
            LevelJustification levelJustification33 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties33 = new PreviousParagraphProperties();
            Indentation indentation34 = new Indentation() { Left = "4320", Hanging = "180" };

            previousParagraphProperties33.Append(indentation34);

            NumberingSymbolRunProperties numberingSymbolRunProperties33 = new NumberingSymbolRunProperties();
            RunFonts runFonts39 = new RunFonts() { ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties33.Append(runFonts39);

            level33.Append(startNumberingValue33);
            level33.Append(numberingFormat33);
            level33.Append(levelText33);
            level33.Append(levelJustification33);
            level33.Append(previousParagraphProperties33);
            level33.Append(numberingSymbolRunProperties33);

            Level level34 = new Level() { LevelIndex = 6, TemplateCode = "0419000F", Tentative = true };
            StartNumberingValue startNumberingValue34 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat34 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText34 = new LevelText() { Val = "%7." };
            LevelJustification levelJustification34 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties34 = new PreviousParagraphProperties();
            Indentation indentation35 = new Indentation() { Left = "5040", Hanging = "360" };

            previousParagraphProperties34.Append(indentation35);

            NumberingSymbolRunProperties numberingSymbolRunProperties34 = new NumberingSymbolRunProperties();
            RunFonts runFonts40 = new RunFonts() { ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties34.Append(runFonts40);

            level34.Append(startNumberingValue34);
            level34.Append(numberingFormat34);
            level34.Append(levelText34);
            level34.Append(levelJustification34);
            level34.Append(previousParagraphProperties34);
            level34.Append(numberingSymbolRunProperties34);

            Level level35 = new Level() { LevelIndex = 7, TemplateCode = "04190019", Tentative = true };
            StartNumberingValue startNumberingValue35 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat35 = new NumberingFormat() { Val = NumberFormatValues.LowerLetter };
            LevelText levelText35 = new LevelText() { Val = "%8." };
            LevelJustification levelJustification35 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties35 = new PreviousParagraphProperties();
            Indentation indentation36 = new Indentation() { Left = "5760", Hanging = "360" };

            previousParagraphProperties35.Append(indentation36);

            NumberingSymbolRunProperties numberingSymbolRunProperties35 = new NumberingSymbolRunProperties();
            RunFonts runFonts41 = new RunFonts() { ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties35.Append(runFonts41);

            level35.Append(startNumberingValue35);
            level35.Append(numberingFormat35);
            level35.Append(levelText35);
            level35.Append(levelJustification35);
            level35.Append(previousParagraphProperties35);
            level35.Append(numberingSymbolRunProperties35);

            Level level36 = new Level() { LevelIndex = 8, TemplateCode = "0419001B", Tentative = true };
            StartNumberingValue startNumberingValue36 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat36 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText36 = new LevelText() { Val = "%9." };
            LevelJustification levelJustification36 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties36 = new PreviousParagraphProperties();
            Indentation indentation37 = new Indentation() { Left = "6480", Hanging = "180" };

            previousParagraphProperties36.Append(indentation37);

            NumberingSymbolRunProperties numberingSymbolRunProperties36 = new NumberingSymbolRunProperties();
            RunFonts runFonts42 = new RunFonts() { ComplexScript = "Times New Roman" };

            numberingSymbolRunProperties36.Append(runFonts42);

            level36.Append(startNumberingValue36);
            level36.Append(numberingFormat36);
            level36.Append(levelText36);
            level36.Append(levelJustification36);
            level36.Append(previousParagraphProperties36);
            level36.Append(numberingSymbolRunProperties36);

            abstractNum4.Append(nsid4);
            abstractNum4.Append(multiLevelType4);
            abstractNum4.Append(templateCode4);
            abstractNum4.Append(level28);
            abstractNum4.Append(level29);
            abstractNum4.Append(level30);
            abstractNum4.Append(level31);
            abstractNum4.Append(level32);
            abstractNum4.Append(level33);
            abstractNum4.Append(level34);
            abstractNum4.Append(level35);
            abstractNum4.Append(level36);

            NumberingInstance numberingInstance1 = new NumberingInstance() { NumberID = 1 };
            AbstractNumId abstractNumId1 = new AbstractNumId() { Val = 2 };

            numberingInstance1.Append(abstractNumId1);

            NumberingInstance numberingInstance2 = new NumberingInstance() { NumberID = 2 };
            AbstractNumId abstractNumId2 = new AbstractNumId() { Val = 3 };

            numberingInstance2.Append(abstractNumId2);

            NumberingInstance numberingInstance3 = new NumberingInstance() { NumberID = 3 };
            AbstractNumId abstractNumId3 = new AbstractNumId() { Val = 0 };

            numberingInstance3.Append(abstractNumId3);

            NumberingInstance numberingInstance4 = new NumberingInstance() { NumberID = 4 };
            AbstractNumId abstractNumId4 = new AbstractNumId() { Val = 1 };

            numberingInstance4.Append(abstractNumId4);

            numbering1.Append(abstractNum1);
            numbering1.Append(abstractNum2);
            numbering1.Append(abstractNum3);
            numbering1.Append(abstractNum4);
            numbering1.Append(numberingInstance1);
            numbering1.Append(numberingInstance2);
            numbering1.Append(numberingInstance3);
            numbering1.Append(numberingInstance4);

            numberingDefinitionsPart1.Numbering = numbering1;
        }

        // Generates content of themePart1.
        private void GenerateThemePart1Content(ThemePart themePart1)
        {
            A.Theme theme1 = new A.Theme() { Name = "Office Theme" };
            theme1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.ThemeElements themeElements1 = new A.ThemeElements();

            A.ColorScheme colorScheme1 = new A.ColorScheme() { Name = "Office" };

            A.Dark1Color dark1Color1 = new A.Dark1Color();
            A.SystemColor systemColor1 = new A.SystemColor() { Val = A.SystemColorValues.WindowText, LastColor = "000000" };

            dark1Color1.Append(systemColor1);

            A.Light1Color light1Color1 = new A.Light1Color();
            A.SystemColor systemColor2 = new A.SystemColor() { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" };

            light1Color1.Append(systemColor2);

            A.Dark2Color dark2Color1 = new A.Dark2Color();
            A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "1F497D" };

            dark2Color1.Append(rgbColorModelHex1);

            A.Light2Color light2Color1 = new A.Light2Color();
            A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "EEECE1" };

            light2Color1.Append(rgbColorModelHex2);

            A.Accent1Color accent1Color1 = new A.Accent1Color();
            A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex() { Val = "4F81BD" };

            accent1Color1.Append(rgbColorModelHex3);

            A.Accent2Color accent2Color1 = new A.Accent2Color();
            A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex() { Val = "C0504D" };

            accent2Color1.Append(rgbColorModelHex4);

            A.Accent3Color accent3Color1 = new A.Accent3Color();
            A.RgbColorModelHex rgbColorModelHex5 = new A.RgbColorModelHex() { Val = "9BBB59" };

            accent3Color1.Append(rgbColorModelHex5);

            A.Accent4Color accent4Color1 = new A.Accent4Color();
            A.RgbColorModelHex rgbColorModelHex6 = new A.RgbColorModelHex() { Val = "8064A2" };

            accent4Color1.Append(rgbColorModelHex6);

            A.Accent5Color accent5Color1 = new A.Accent5Color();
            A.RgbColorModelHex rgbColorModelHex7 = new A.RgbColorModelHex() { Val = "4BACC6" };

            accent5Color1.Append(rgbColorModelHex7);

            A.Accent6Color accent6Color1 = new A.Accent6Color();
            A.RgbColorModelHex rgbColorModelHex8 = new A.RgbColorModelHex() { Val = "F79646" };

            accent6Color1.Append(rgbColorModelHex8);

            A.Hyperlink hyperlink1 = new A.Hyperlink();
            A.RgbColorModelHex rgbColorModelHex9 = new A.RgbColorModelHex() { Val = "0000FF" };

            hyperlink1.Append(rgbColorModelHex9);

            A.FollowedHyperlinkColor followedHyperlinkColor1 = new A.FollowedHyperlinkColor();
            A.RgbColorModelHex rgbColorModelHex10 = new A.RgbColorModelHex() { Val = "800080" };

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

            A.FontScheme fontScheme1 = new A.FontScheme() { Name = "Office" };

            A.MajorFont majorFont1 = new A.MajorFont();
            A.LatinFont latinFont1 = new A.LatinFont() { Typeface = "Cambria" };
            A.EastAsianFont eastAsianFont1 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont1 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont1 = new A.SupplementalFont() { Script = "Jpan", Typeface = "ＭＳ ゴシック" };
            A.SupplementalFont supplementalFont2 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont3 = new A.SupplementalFont() { Script = "Hans", Typeface = "宋体" };
            A.SupplementalFont supplementalFont4 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont5 = new A.SupplementalFont() { Script = "Arab", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont6 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont7 = new A.SupplementalFont() { Script = "Thai", Typeface = "Angsana New" };
            A.SupplementalFont supplementalFont8 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont9 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont10 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont11 = new A.SupplementalFont() { Script = "Khmr", Typeface = "MoolBoran" };
            A.SupplementalFont supplementalFont12 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont13 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont14 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont15 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont16 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont17 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont18 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont19 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont20 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont21 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont22 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont23 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont24 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont25 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont26 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont27 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont28 = new A.SupplementalFont() { Script = "Viet", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont29 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };

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
            A.LatinFont latinFont2 = new A.LatinFont() { Typeface = "Calibri" };
            A.EastAsianFont eastAsianFont2 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont2 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont30 = new A.SupplementalFont() { Script = "Jpan", Typeface = "ＭＳ 明朝" };
            A.SupplementalFont supplementalFont31 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont32 = new A.SupplementalFont() { Script = "Hans", Typeface = "宋体" };
            A.SupplementalFont supplementalFont33 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont34 = new A.SupplementalFont() { Script = "Arab", Typeface = "Arial" };
            A.SupplementalFont supplementalFont35 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Arial" };
            A.SupplementalFont supplementalFont36 = new A.SupplementalFont() { Script = "Thai", Typeface = "Cordia New" };
            A.SupplementalFont supplementalFont37 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont38 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont39 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont40 = new A.SupplementalFont() { Script = "Khmr", Typeface = "DaunPenh" };
            A.SupplementalFont supplementalFont41 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont42 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont43 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont44 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont45 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont46 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont47 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont48 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont49 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont50 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont51 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont52 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont53 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont54 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont55 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont56 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont57 = new A.SupplementalFont() { Script = "Viet", Typeface = "Arial" };
            A.SupplementalFont supplementalFont58 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };

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

            A.FormatScheme formatScheme1 = new A.FormatScheme() { Name = "Office" };

            A.FillStyleList fillStyleList1 = new A.FillStyleList();

            A.SolidFill solidFill1 = new A.SolidFill();
            A.SchemeColor schemeColor1 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill1.Append(schemeColor1);

            A.GradientFill gradientFill1 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList1 = new A.GradientStopList();

            A.GradientStop gradientStop1 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor2 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint1 = new A.Tint() { Val = 50000 };
            A.SaturationModulation saturationModulation1 = new A.SaturationModulation() { Val = 300000 };

            schemeColor2.Append(tint1);
            schemeColor2.Append(saturationModulation1);

            gradientStop1.Append(schemeColor2);

            A.GradientStop gradientStop2 = new A.GradientStop() { Position = 35000 };

            A.SchemeColor schemeColor3 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint2 = new A.Tint() { Val = 37000 };
            A.SaturationModulation saturationModulation2 = new A.SaturationModulation() { Val = 300000 };

            schemeColor3.Append(tint2);
            schemeColor3.Append(saturationModulation2);

            gradientStop2.Append(schemeColor3);

            A.GradientStop gradientStop3 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor4 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint3 = new A.Tint() { Val = 15000 };
            A.SaturationModulation saturationModulation3 = new A.SaturationModulation() { Val = 350000 };

            schemeColor4.Append(tint3);
            schemeColor4.Append(saturationModulation3);

            gradientStop3.Append(schemeColor4);

            gradientStopList1.Append(gradientStop1);
            gradientStopList1.Append(gradientStop2);
            gradientStopList1.Append(gradientStop3);
            A.LinearGradientFill linearGradientFill1 = new A.LinearGradientFill() { Angle = 16200000, Scaled = true };

            gradientFill1.Append(gradientStopList1);
            gradientFill1.Append(linearGradientFill1);

            A.GradientFill gradientFill2 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList2 = new A.GradientStopList();

            A.GradientStop gradientStop4 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor5 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade1 = new A.Shade() { Val = 51000 };
            A.SaturationModulation saturationModulation4 = new A.SaturationModulation() { Val = 130000 };

            schemeColor5.Append(shade1);
            schemeColor5.Append(saturationModulation4);

            gradientStop4.Append(schemeColor5);

            A.GradientStop gradientStop5 = new A.GradientStop() { Position = 80000 };

            A.SchemeColor schemeColor6 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade2 = new A.Shade() { Val = 93000 };
            A.SaturationModulation saturationModulation5 = new A.SaturationModulation() { Val = 130000 };

            schemeColor6.Append(shade2);
            schemeColor6.Append(saturationModulation5);

            gradientStop5.Append(schemeColor6);

            A.GradientStop gradientStop6 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor7 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade3 = new A.Shade() { Val = 94000 };
            A.SaturationModulation saturationModulation6 = new A.SaturationModulation() { Val = 135000 };

            schemeColor7.Append(shade3);
            schemeColor7.Append(saturationModulation6);

            gradientStop6.Append(schemeColor7);

            gradientStopList2.Append(gradientStop4);
            gradientStopList2.Append(gradientStop5);
            gradientStopList2.Append(gradientStop6);
            A.LinearGradientFill linearGradientFill2 = new A.LinearGradientFill() { Angle = 16200000, Scaled = false };

            gradientFill2.Append(gradientStopList2);
            gradientFill2.Append(linearGradientFill2);

            fillStyleList1.Append(solidFill1);
            fillStyleList1.Append(gradientFill1);
            fillStyleList1.Append(gradientFill2);

            A.LineStyleList lineStyleList1 = new A.LineStyleList();

            A.Outline outline1 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill2 = new A.SolidFill();

            A.SchemeColor schemeColor8 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade4 = new A.Shade() { Val = 95000 };
            A.SaturationModulation saturationModulation7 = new A.SaturationModulation() { Val = 105000 };

            schemeColor8.Append(shade4);
            schemeColor8.Append(saturationModulation7);

            solidFill2.Append(schemeColor8);
            A.PresetDash presetDash1 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            outline1.Append(solidFill2);
            outline1.Append(presetDash1);

            A.Outline outline2 = new A.Outline() { Width = 25400, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill3 = new A.SolidFill();
            A.SchemeColor schemeColor9 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill3.Append(schemeColor9);
            A.PresetDash presetDash2 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            outline2.Append(solidFill3);
            outline2.Append(presetDash2);

            A.Outline outline3 = new A.Outline() { Width = 38100, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill4 = new A.SolidFill();
            A.SchemeColor schemeColor10 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill4.Append(schemeColor10);
            A.PresetDash presetDash3 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            outline3.Append(solidFill4);
            outline3.Append(presetDash3);

            lineStyleList1.Append(outline1);
            lineStyleList1.Append(outline2);
            lineStyleList1.Append(outline3);

            A.EffectStyleList effectStyleList1 = new A.EffectStyleList();

            A.EffectStyle effectStyle1 = new A.EffectStyle();

            A.EffectList effectList1 = new A.EffectList();

            A.OuterShadow outerShadow1 = new A.OuterShadow() { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex11 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha1 = new A.Alpha() { Val = 38000 };

            rgbColorModelHex11.Append(alpha1);

            outerShadow1.Append(rgbColorModelHex11);

            effectList1.Append(outerShadow1);

            effectStyle1.Append(effectList1);

            A.EffectStyle effectStyle2 = new A.EffectStyle();

            A.EffectList effectList2 = new A.EffectList();

            A.OuterShadow outerShadow2 = new A.OuterShadow() { BlurRadius = 40000L, Distance = 23000L, Direction = 5400000, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex12 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha2 = new A.Alpha() { Val = 35000 };

            rgbColorModelHex12.Append(alpha2);

            outerShadow2.Append(rgbColorModelHex12);

            effectList2.Append(outerShadow2);

            effectStyle2.Append(effectList2);

            A.EffectStyle effectStyle3 = new A.EffectStyle();

            A.EffectList effectList3 = new A.EffectList();

            A.OuterShadow outerShadow3 = new A.OuterShadow() { BlurRadius = 40000L, Distance = 23000L, Direction = 5400000, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex13 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha3 = new A.Alpha() { Val = 35000 };

            rgbColorModelHex13.Append(alpha3);

            outerShadow3.Append(rgbColorModelHex13);

            effectList3.Append(outerShadow3);

            A.Scene3DType scene3DType1 = new A.Scene3DType();

            A.Camera camera1 = new A.Camera() { Preset = A.PresetCameraValues.OrthographicFront };
            A.Rotation rotation1 = new A.Rotation() { Latitude = 0, Longitude = 0, Revolution = 0 };

            camera1.Append(rotation1);

            A.LightRig lightRig1 = new A.LightRig() { Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };
            A.Rotation rotation2 = new A.Rotation() { Latitude = 0, Longitude = 0, Revolution = 1200000 };

            lightRig1.Append(rotation2);

            scene3DType1.Append(camera1);
            scene3DType1.Append(lightRig1);

            A.Shape3DType shape3DType1 = new A.Shape3DType();
            A.BevelTop bevelTop1 = new A.BevelTop() { Width = 63500L, Height = 25400L };

            shape3DType1.Append(bevelTop1);

            effectStyle3.Append(effectList3);
            effectStyle3.Append(scene3DType1);
            effectStyle3.Append(shape3DType1);

            effectStyleList1.Append(effectStyle1);
            effectStyleList1.Append(effectStyle2);
            effectStyleList1.Append(effectStyle3);

            A.BackgroundFillStyleList backgroundFillStyleList1 = new A.BackgroundFillStyleList();

            A.SolidFill solidFill5 = new A.SolidFill();
            A.SchemeColor schemeColor11 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill5.Append(schemeColor11);

            A.GradientFill gradientFill3 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList3 = new A.GradientStopList();

            A.GradientStop gradientStop7 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor12 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint4 = new A.Tint() { Val = 40000 };
            A.SaturationModulation saturationModulation8 = new A.SaturationModulation() { Val = 350000 };

            schemeColor12.Append(tint4);
            schemeColor12.Append(saturationModulation8);

            gradientStop7.Append(schemeColor12);

            A.GradientStop gradientStop8 = new A.GradientStop() { Position = 40000 };

            A.SchemeColor schemeColor13 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint5 = new A.Tint() { Val = 45000 };
            A.Shade shade5 = new A.Shade() { Val = 99000 };
            A.SaturationModulation saturationModulation9 = new A.SaturationModulation() { Val = 350000 };

            schemeColor13.Append(tint5);
            schemeColor13.Append(shade5);
            schemeColor13.Append(saturationModulation9);

            gradientStop8.Append(schemeColor13);

            A.GradientStop gradientStop9 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor14 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade6 = new A.Shade() { Val = 20000 };
            A.SaturationModulation saturationModulation10 = new A.SaturationModulation() { Val = 255000 };

            schemeColor14.Append(shade6);
            schemeColor14.Append(saturationModulation10);

            gradientStop9.Append(schemeColor14);

            gradientStopList3.Append(gradientStop7);
            gradientStopList3.Append(gradientStop8);
            gradientStopList3.Append(gradientStop9);

            A.PathGradientFill pathGradientFill1 = new A.PathGradientFill() { Path = A.PathShadeValues.Circle };
            A.FillToRectangle fillToRectangle1 = new A.FillToRectangle() { Left = 50000, Top = -80000, Right = 50000, Bottom = 180000 };

            pathGradientFill1.Append(fillToRectangle1);

            gradientFill3.Append(gradientStopList3);
            gradientFill3.Append(pathGradientFill1);

            A.GradientFill gradientFill4 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList4 = new A.GradientStopList();

            A.GradientStop gradientStop10 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor15 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint6 = new A.Tint() { Val = 80000 };
            A.SaturationModulation saturationModulation11 = new A.SaturationModulation() { Val = 300000 };

            schemeColor15.Append(tint6);
            schemeColor15.Append(saturationModulation11);

            gradientStop10.Append(schemeColor15);

            A.GradientStop gradientStop11 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor16 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade7 = new A.Shade() { Val = 30000 };
            A.SaturationModulation saturationModulation12 = new A.SaturationModulation() { Val = 200000 };

            schemeColor16.Append(shade7);
            schemeColor16.Append(saturationModulation12);

            gradientStop11.Append(schemeColor16);

            gradientStopList4.Append(gradientStop10);
            gradientStopList4.Append(gradientStop11);

            A.PathGradientFill pathGradientFill2 = new A.PathGradientFill() { Path = A.PathShadeValues.Circle };
            A.FillToRectangle fillToRectangle2 = new A.FillToRectangle() { Left = 50000, Top = 50000, Right = 50000, Bottom = 50000 };

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

        // Generates content of fontTablePart1.
        private void GenerateFontTablePart1Content(FontTablePart fontTablePart1)
        {
            Fonts fonts1 = new Fonts();
            fonts1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            fonts1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            Font font1 = new Font() { Name = "Times New Roman" };
            Panose1Number panose1Number1 = new Panose1Number() { Val = "02020603050405020304" };
            FontCharSet fontCharSet1 = new FontCharSet() { Val = "CC" };
            FontFamily fontFamily1 = new FontFamily() { Val = FontFamilyValues.Roman };
            Pitch pitch1 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature1 = new FontSignature() { UnicodeSignature0 = "E0002AFF", UnicodeSignature1 = "C0007841", UnicodeSignature2 = "00000009", UnicodeSignature3 = "00000000", CodePageSignature0 = "000001FF", CodePageSignature1 = "00000000" };

            font1.Append(panose1Number1);
            font1.Append(fontCharSet1);
            font1.Append(fontFamily1);
            font1.Append(pitch1);
            font1.Append(fontSignature1);

            Font font2 = new Font() { Name = "Calibri" };
            Panose1Number panose1Number2 = new Panose1Number() { Val = "020F0502020204030204" };
            FontCharSet fontCharSet2 = new FontCharSet() { Val = "CC" };
            FontFamily fontFamily2 = new FontFamily() { Val = FontFamilyValues.Swiss };
            Pitch pitch2 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature2 = new FontSignature() { UnicodeSignature0 = "E10002FF", UnicodeSignature1 = "4000ACFF", UnicodeSignature2 = "00000009", UnicodeSignature3 = "00000000", CodePageSignature0 = "0000019F", CodePageSignature1 = "00000000" };

            font2.Append(panose1Number2);
            font2.Append(fontCharSet2);
            font2.Append(fontFamily2);
            font2.Append(pitch2);
            font2.Append(fontSignature2);

            Font font3 = new Font() { Name = "Tahoma" };
            Panose1Number panose1Number3 = new Panose1Number() { Val = "020B0604030504040204" };
            FontCharSet fontCharSet3 = new FontCharSet() { Val = "CC" };
            FontFamily fontFamily3 = new FontFamily() { Val = FontFamilyValues.Swiss };
            Pitch pitch3 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature3 = new FontSignature() { UnicodeSignature0 = "E1002EFF", UnicodeSignature1 = "C000605B", UnicodeSignature2 = "00000029", UnicodeSignature3 = "00000000", CodePageSignature0 = "000101FF", CodePageSignature1 = "00000000" };

            font3.Append(panose1Number3);
            font3.Append(fontCharSet3);
            font3.Append(fontFamily3);
            font3.Append(pitch3);
            font3.Append(fontSignature3);

            Font font4 = new Font() { Name = "Cambria" };
            Panose1Number panose1Number4 = new Panose1Number() { Val = "02040503050406030204" };
            FontCharSet fontCharSet4 = new FontCharSet() { Val = "CC" };
            FontFamily fontFamily4 = new FontFamily() { Val = FontFamilyValues.Roman };
            Pitch pitch4 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature4 = new FontSignature() { UnicodeSignature0 = "E00002FF", UnicodeSignature1 = "400004FF", UnicodeSignature2 = "00000000", UnicodeSignature3 = "00000000", CodePageSignature0 = "0000019F", CodePageSignature1 = "00000000" };

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

        // Generates content of webSettingsPart1.
        private void GenerateWebSettingsPart1Content(WebSettingsPart webSettingsPart1)
        {
            WebSettings webSettings1 = new WebSettings();
            webSettings1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            webSettings1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            OptimizeForBrowser optimizeForBrowser1 = new OptimizeForBrowser();

            webSettings1.Append(optimizeForBrowser1);

            webSettingsPart1.WebSettings = webSettings1;
        }

        private void SetPackageProperties(OpenXmlPackage document)
        {
            document.PackageProperties.Creator = "Админ";
            document.PackageProperties.Title = "Министерство образования и науки РФ ";
            document.PackageProperties.Subject = "";
            document.PackageProperties.Keywords = "";
            document.PackageProperties.Description = "";
            document.PackageProperties.Revision = "3";
            document.PackageProperties.Created = System.Xml.XmlConvert.ToDateTime("2017-12-06T19:30:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2017-12-20T21:53:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.LastModifiedBy = "Grisha";
        }


    }
}
