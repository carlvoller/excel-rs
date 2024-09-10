use std::io::{Result, Seek, Write};

use zip::{write::SimpleFileOptions, ZipWriter};

pub struct XlsxFormatter<W: Write + Seek> {
    pub zip_writer: ZipWriter<W>,
}

impl<W: Write + Seek> XlsxFormatter<W> {
    pub fn new(zip_writer: ZipWriter<W>) -> Self {
        XlsxFormatter { zip_writer }
    }

    // pub fn write_sheet(&mut self, sheet: Sheet) -> Result<()> {
    //     let sheet_id = sheet.id;
    //     let sheet_buf = sheet.close().ok().unwrap();

    //     let options = SimpleFileOptions::default()
    //         .compression_method(zip::CompressionMethod::Deflated)
    //         .compression_level(Some(1));
    //     self.zip_writer
    //         .start_file(format!("xl/worksheets/sheet{}.xml", sheet_id), options)?;

    //     self.zip_writer.write_all(&sheet_buf)?;

    //     Ok(())
    // }

    pub fn finish(mut self, num_of_sheets: u16) -> Result<W> {
        let options = SimpleFileOptions::default();
        self.write_content_type(&options, num_of_sheets)?;
        self.write_rels(&options)?;
        self.write_doc_props(&options)?;
        self.write_styles(&options)?;
        self.write_shared_strings(&options)?;
        self.write_work_book(&options, num_of_sheets)?;
        self.write_calc_chain(&options)?;
        self.write_xl_rels(&options, num_of_sheets)?;
        self.write_theme(&options)?;
        let result = self.zip_writer.finish()?;
        Ok(result)
    }

    fn write_content_type(
        &mut self,
        options: &SimpleFileOptions,
        num_of_sheets: u16,
    ) -> Result<()> {
        self.zip_writer
            .start_file("[Content_Types].xml", *options)?;
        write!(
            self.zip_writer,
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><Default Extension="xml" ContentType="application/xml"/><Default Extension="bin" ContentType="application/vnd.ms-excel.sheet.binary.macroEnabled.main"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>"#
        )?;
        for i in 0..num_of_sheets {
            writeln!(self.zip_writer, "<Override PartName=\"/xl/worksheets/sheet{}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>", i + 1)?;
        }
        write!(
            self.zip_writer,
            r#"<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>"#
        )
    }

    fn write_rels(&mut self, options: &SimpleFileOptions) -> Result<()> {
        self.zip_writer.start_file("_rels/.rels", *options)?;
        write!(
            self.zip_writer,
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#
        )
    }

    fn write_doc_props(&mut self, options: &SimpleFileOptions) -> Result<()> {
        self.zip_writer.start_file("docProps/app.xml", *options)?;
        write!(
            self.zip_writer,
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
                    xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
            <Application>SheetJS</Application>
            <HeadingPairs>
                <vt:vector size="2" baseType="variant">
                    <vt:variant>
                        <vt:lpstr>Worksheets</vt:lpstr>
                    </vt:variant>
                    <vt:variant>
                        <vt:i4>1</vt:i4>
                    </vt:variant>
                </vt:vector>
            </HeadingPairs>
            <TitlesOfParts>
                <vt:vector size="1" baseType="lpstr">
                <vt:lpstr>SheetJS</vt:lpstr>
                </vt:vector>
            </TitlesOfParts>
        </Properties>"#
        )?;
        self.zip_writer.start_file("docProps/core.xml", *options)?;
        write!(
            self.zip_writer,
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                           xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"
                           xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"/>"#
        )
    }

    fn write_styles(&mut self, options: &SimpleFileOptions) -> Result<()> {
        self.zip_writer.start_file("xl/styles.xml", *options)?;
        write!(
            self.zip_writer,
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                    xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
            <fonts count="1">
                <font>
                    <sz val="12"/>
                    <color theme="1"/>
                    <name val="Calibri"/>
                    <family val="2"/>
                    <scheme val="minor"/>
                </font>
            </fonts>
            <fills count="2">
                <fill>
                    <patternFill patternType="none"/>
                </fill>
                <fill>
                    <patternFill patternType="gray125"/>
                </fill>
            </fills>
            <borders count="1">
                <border>
                    <left/>
                    <right/>
                    <top/>
                    <bottom/>
                    <diagonal/>
                </border>
            </borders>
        <cellStyleXfs count="1">
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
        </cellStyleXfs>
        <cellXfs count="1">
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
        </cellXfs>
        <cellStyles count="1">
            <cellStyle name="Normal" xfId="0" builtinId="0"/>
        </cellStyles>
        <dxfs count="0"/>
        <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleMedium4"/>
    </styleSheet>"#
        )
    }

    fn write_shared_strings(&mut self, options: &SimpleFileOptions) -> Result<()> {
        self.zip_writer
            .start_file("xl/sharedStrings.xml", *options)?;
        write!(
            self.zip_writer,
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"></sst>"#
        )
    }

    fn write_work_book(&mut self, options: &SimpleFileOptions, num_of_sheets: u16) -> Result<()> {
        self.zip_writer.start_file("xl/workbook.xml", *options)?;
        write!(
            self.zip_writer,
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <workbookPr date1904="false"/>
            <sheets>
    "#
        )?;
        for i in 0..num_of_sheets {
            writeln!(
                self.zip_writer,
                "<sheet name=\"Sheet {}\" sheetId=\"{}\" r:id=\"rId{}\"/>",
                i + 1,
                i + 1,
                i + 3
            )?;
        }
        write!(
            self.zip_writer,
            r#"
        </sheets>
    </workbook>
    "#
        )
    }

    fn write_calc_chain(&mut self, options: &SimpleFileOptions) -> Result<()> {
        self.zip_writer.start_file("xl/calcChain.xml", *options)?;
        write!(
            self.zip_writer,
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></calcChain>"#
        )
    }

    fn write_xl_rels(&mut self, options: &SimpleFileOptions, num_of_sheets: u16) -> Result<()> {
        self.zip_writer
            .start_file("xl/_rels/workbook.xml.rels", *options)?;
        write!(
            self.zip_writer,
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
        <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
                "#
        )?;
        let mut last_rid = 2;
        for i in 0..num_of_sheets {
            writeln!(
                self.zip_writer,
                "<Relationship Id=\"rId{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet{}.xml\"/>", i + 3, i + 1
            )?;
            last_rid = i + 3;
        }
        write!(
            self.zip_writer,
            r#"
            <Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
        </Relationships>"#,
            last_rid + 1
        )
    }

    fn write_theme(&mut self, options: &SimpleFileOptions) -> Result<()> {
        self.zip_writer
            .start_file("xl/theme/theme1.xml", *options)?;
        write!(
            self.zip_writer,
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
            <a:themeElements>
                <a:clrScheme name="Office">
                    <a:dk1>
                        <a:sysClr val="windowText" lastClr="000000"/>
                    </a:dk1>
                    <a:lt1>
                        <a:sysClr val="window" lastClr="FFFFFF"/>
                    </a:lt1>
                    <a:dk2>
                        <a:srgbClr val="1F497D"/>
                    </a:dk2>
                    <a:lt2>
                        <a:srgbClr val="EEECE1"/>
                    </a:lt2>
                    <a:accent1>
                        <a:srgbClr val="4F81BD"/>
                    </a:accent1>
                    <a:accent2>
                        <a:srgbClr val="C0504D"/>
                    </a:accent2>
                    <a:accent3>
                        <a:srgbClr val="9BBB59"/>
                    </a:accent3>
                    <a:accent4>
                        <a:srgbClr val="8064A2"/>
                    </a:accent4>
                    <a:accent5>
                        <a:srgbClr val="4BACC6"/>
                    </a:accent5>
                    <a:accent6>
                        <a:srgbClr val="F79646"/>
                    </a:accent6>
                    <a:hlink>
                        <a:srgbClr val="0000FF"/>
                    </a:hlink>
                    <a:folHlink>
                        <a:srgbClr val="800080"/>
                    </a:folHlink>
                </a:clrScheme>
                <a:fontScheme name="Office">
                    <a:majorFont>
                        <a:latin typeface="Cambria"/>
                        <a:ea typeface=""/>
                        <a:cs typeface=""/>
                        <a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/>
                        <a:font script="Hang" typeface="맑은 고딕"/>
                        <a:font script="Hans" typeface="宋体"/>
                        <a:font script="Hant" typeface="新細明體"/>
                        <a:font script="Arab" typeface="Times New Roman"/>
                        <a:font script="Hebr" typeface="Times New Roman"/>
                        <a:font script="Thai" typeface="Tahoma"/>
                        <a:font script="Ethi" typeface="Nyala"/>
                        <a:font script="Beng" typeface="Vrinda"/>
                        <a:font script="Gujr" typeface="Shruti"/>
                        <a:font script="Khmr" typeface="MoolBoran"/>
                        <a:font script="Knda" typeface="Tunga"/>
                        <a:font script="Guru" typeface="Raavi"/>
                        <a:font script="Cans" typeface="Euphemia"/>
                        <a:font script="Cher" typeface="Plantagenet Cherokee"/>
                        <a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
                        <a:font script="Tibt" typeface="Microsoft Himalaya"/>
                        <a:font script="Thaa" typeface="MV Boli"/>
                        <a:font script="Deva" typeface="Mangal"/>
                        <a:font script="Telu" typeface="Gautami"/>
                        <a:font script="Taml" typeface="Latha"/>
                        <a:font script="Syrc" typeface="Estrangelo Edessa"/>
                        <a:font script="Orya" typeface="Kalinga"/>
                        <a:font script="Mlym" typeface="Kartika"/>
                        <a:font script="Laoo" typeface="DokChampa"/>
                        <a:font script="Sinh" typeface="Iskoola Pota"/>
                        <a:font script="Mong" typeface="Mongolian Baiti"/>
                        <a:font script="Viet" typeface="Times New Roman"/>
                        <a:font script="Uigh" typeface="Microsoft Uighur"/>
                        <a:font script="Geor" typeface="Sylfaen"/>
                    </a:majorFont>
                    <a:minorFont>
                        <a:latin typeface="Calibri"/>
                        <a:ea typeface=""/>
                        <a:cs typeface=""/>
                        <a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/>
                        <a:font script="Hang" typeface="맑은 고딕"/>
                        <a:font script="Hans" typeface="宋体"/>
                        <a:font script="Hant" typeface="新細明體"/>
                        <a:font script="Arab" typeface="Arial"/>
                        <a:font script="Hebr" typeface="Arial"/>
                        <a:font script="Thai" typeface="Tahoma"/>
                        <a:font script="Ethi" typeface="Nyala"/>
                        <a:font script="Beng" typeface="Vrinda"/>
                        <a:font script="Gujr" typeface="Shruti"/>
                        <a:font script="Khmr" typeface="DaunPenh"/>
                        <a:font script="Knda" typeface="Tunga"/>
                        <a:font script="Guru" typeface="Raavi"/>
                        <a:font script="Cans" typeface="Euphemia"/>
                        <a:font script="Cher" typeface="Plantagenet Cherokee"/>
                        <a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
                        <a:font script="Tibt" typeface="Microsoft Himalaya"/>
                        <a:font script="Thaa" typeface="MV Boli"/>
                        <a:font script="Deva" typeface="Mangal"/>
                        <a:font script="Telu" typeface="Gautami"/>
                        <a:font script="Taml" typeface="Latha"/>
                        <a:font script="Syrc" typeface="Estrangelo Edessa"/>
                        <a:font script="Orya" typeface="Kalinga"/>
                        <a:font script="Mlym" typeface="Kartika"/>
                        <a:font script="Laoo" typeface="DokChampa"/>
                        <a:font script="Sinh" typeface="Iskoola Pota"/>
                        <a:font script="Mong" typeface="Mongolian Baiti"/>
                        <a:font script="Viet" typeface="Arial"/>
                        <a:font script="Uigh" typeface="Microsoft Uighur"/>
                        <a:font script="Geor" typeface="Sylfaen"/>
                    </a:minorFont>
                </a:fontScheme>
                <a:fmtScheme name="Office">
                    <a:fillStyleLst>
                        <a:solidFill>
                            <a:schemeClr val="phClr"/>
                        </a:solidFill>
                        <a:gradFill rotWithShape="1">
                            <a:gsLst>
                                <a:gs pos="0">
                                    <a:schemeClr val="phClr">
                                        <a:tint val="50000"/>
                                        <a:satMod val="300000"/>
                                    </a:schemeClr>
                                </a:gs>
                                <a:gs pos="35000">
                                    <a:schemeClr val="phClr">
                                        <a:tint val="37000"/>
                                        <a:satMod val="300000"/>
                                    </a:schemeClr>
                                </a:gs>
                                <a:gs pos="100000">
                                    <a:schemeClr val="phClr">
                                        <a:tint val="15000"/>
                                        <a:satMod val="350000"/>
                                    </a:schemeClr>
                                </a:gs>
                            </a:gsLst>
                            <a:lin ang="16200000" scaled="1"/>
                        </a:gradFill>
                        <a:gradFill rotWithShape="1">
                            <a:gsLst>
                                <a:gs pos="0">
                                    <a:schemeClr val="phClr">
                                        <a:tint val="100000"/>
                                        <a:shade val="100000"/>
                                        <a:satMod val="130000"/>
                                    </a:schemeClr>
                                </a:gs>
                                <a:gs pos="100000">
                                    <a:schemeClr val="phClr">
                                        <a:tint val="50000"/>
                                        <a:shade val="100000"/>
                                        <a:satMod val="350000"/>
                                    </a:schemeClr>
                                </a:gs>
                            </a:gsLst>
                            <a:lin ang="16200000" scaled="0"/>
                        </a:gradFill>
                    </a:fillStyleLst>
                    <a:lnStyleLst>
                        <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
                            <a:solidFill>
                                <a:schemeClr val="phClr">
                                    <a:shade val="95000"/>
                                    <a:satMod val="105000"/>
                                </a:schemeClr>
                            </a:solidFill>
                            <a:prstDash val="solid"/>
                        </a:ln>
                        <a:ln w="25400" cap="flat" cmpd="sng" algn="ctr">
                            <a:solidFill>
                                <a:schemeClr val="phClr"/>
                            </a:solidFill>
                            <a:prstDash val="solid"/>
                        </a:ln>
                        <a:ln w="38100" cap="flat" cmpd="sng" algn="ctr">
                            <a:solidFill>
                                <a:schemeClr val="phClr"/>
                            </a:solidFill>
                            <a:prstDash val="solid"/>
                        </a:ln>
                    </a:lnStyleLst>
                    <a:effectStyleLst>
                        <a:effectStyle>
                            <a:effectLst>
                                <a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0">
                                    <a:srgbClr val="000000">
                                        <a:alpha val="38000"/>
                                    </a:srgbClr>
                                </a:outerShdw>
                            </a:effectLst>
                        </a:effectStyle>
                        <a:effectStyle>
                            <a:effectLst>
                                <a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
                                    <a:srgbClr val="000000">
                                        <a:alpha val="35000"/>
                                    </a:srgbClr>
                                </a:outerShdw>
                            </a:effectLst>
                        </a:effectStyle>
                        <a:effectStyle>
                            <a:effectLst>
                                <a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
                                    <a:srgbClr val="000000">
                                        <a:alpha val="35000"/>
                                    </a:srgbClr>
                                </a:outerShdw>
                            </a:effectLst>
                            <a:scene3d>
                                <a:camera prst="orthographicFront">
                                    <a:rot lat="0" lon="0" rev="0"/>
                                </a:camera>
                                <a:lightRig rig="threePt" dir="t">
                                    <a:rot lat="0" lon="0" rev="1200000"/>
                                </a:lightRig>
                            </a:scene3d>
                            <a:sp3d>
                                <a:bevelT w="63500" h="25400"/>
                            </a:sp3d>
                        </a:effectStyle>
                    </a:effectStyleLst>
                    <a:bgFillStyleLst>
                        <a:solidFill>
                            <a:schemeClr val="phClr"/>
                        </a:solidFill>
                        <a:gradFill rotWithShape="1">
                            <a:gsLst>
                                <a:gs pos="0">
                                    <a:schemeClr val="phClr">
                                        <a:tint val="40000"/>
                                        <a:satMod val="350000"/>
                                    </a:schemeClr>
                                </a:gs>
                                <a:gs pos="40000">
                                    <a:schemeClr val="phClr">
                                        <a:tint val="45000"/>
                                        <a:shade val="99000"/>
                                        <a:satMod val="350000"/>
                                    </a:schemeClr>
                                </a:gs>
                                <a:gs pos="100000">
                                    <a:schemeClr val="phClr">
                                        <a:shade val="20000"/>
                                        <a:satMod val="255000"/>
                                    </a:schemeClr>
                                </a:gs>
                            </a:gsLst>
                            <a:path path="circle">
                                <a:fillToRect l="50000" t="-80000" r="50000" b="180000"/>
                            </a:path>
                        </a:gradFill>
                        <a:gradFill rotWithShape="1">
                            <a:gsLst>
                                <a:gs pos="0">
                                    <a:schemeClr val="phClr">
                                        <a:tint val="80000"/>
                                        <a:satMod val="300000"/>
                                    </a:schemeClr>
                                </a:gs>
                                <a:gs pos="100000">
                                    <a:schemeClr val="phClr">
                                        <a:shade val="30000"/>
                                        <a:satMod val="200000"/>
                                    </a:schemeClr>
                                </a:gs>
                            </a:gsLst>
                            <a:path path="circle">
                                <a:fillToRect l="50000" t="50000" r="50000" b="50000"/>
                            </a:path>
                        </a:gradFill>
                    </a:bgFillStyleLst>
                </a:fmtScheme>
            </a:themeElements>
            <a:objectDefaults>
                <a:spDef>
                    <a:spPr/>
                    <a:bodyPr/>
                    <a:lstStyle/>
                    <a:style>
                        <a:lnRef idx="1">
                            <a:schemeClr val="accent1"/>
                        </a:lnRef>
                        <a:fillRef idx="3">
                            <a:schemeClr val="accent1"/>
                        </a:fillRef>
                        <a:effectRef idx="2">
                            <a:schemeClr val="accent1"/>
                        </a:effectRef>
                        <a:fontRef idx="minor">
                            <a:schemeClr val="lt1"/>
                        </a:fontRef>
                    </a:style>
                </a:spDef>
                <a:lnDef>
                    <a:spPr/>
                    <a:bodyPr/>
                    <a:lstStyle/>
                    <a:style>
                        <a:lnRef idx="2">
                            <a:schemeClr val="accent1"/>
                        </a:lnRef>
                        <a:fillRef idx="0">
                            <a:schemeClr val="accent1"/>
                        </a:fillRef>
                        <a:effectRef idx="1">
                            <a:schemeClr val="accent1"/>
                        </a:effectRef>
                        <a:fontRef idx="minor">
                            <a:schemeClr val="tx1"/>
                        </a:fontRef>
                    </a:style>
                </a:lnDef>
            </a:objectDefaults>
            <a:extraClrSchemeLst/>
        </a:theme>"#
        )
    }
}
