describe("Render drawing shape", function () {
  // 构造最小 DOCX 文档，专门覆盖 WordprocessingShape 绘图分支
  const createDrawingShapeDocx = async () => {
    const zip = new JSZip(); // DOCX 压缩包容器
    const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`; // 内容类型定义
    const rootRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`; // 顶层关系定义
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  mc:Ignorable="wps">
  <w:body>
    <w:p>
      <w:r>
        <mc:AlternateContent>
          <mc:Choice Requires="wps">
            <w:drawing>
              <wp:inline>
                <wp:extent cx="1828800" cy="914400"/>
                <wp:docPr id="1" name="Drawing 1"/>
                <a:graphic>
                  <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                    <wps:wsp>
                      <wps:spPr>
                        <a:xfrm>
                          <a:off x="0" y="0"/>
                          <a:ext cx="1828800" cy="914400"/>
                        </a:xfrm>
                        <a:prstGeom prst="roundRect">
                          <a:avLst/>
                        </a:prstGeom>
                        <a:solidFill>
                          <a:srgbClr val="FFF2CC"/>
                        </a:solidFill>
                        <a:ln w="12700">
                          <a:solidFill>
                            <a:srgbClr val="C65911"/>
                          </a:solidFill>
                        </a:ln>
                      </wps:spPr>
                      <wps:bodyPr lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="ctr"/>
                      <wps:txbx>
                        <w:txbxContent>
                          <w:p>
                            <w:r>
                              <w:t>绘图文本</w:t>
                            </w:r>
                          </w:p>
                        </w:txbxContent>
                      </wps:txbx>
                    </wps:wsp>
                  </a:graphicData>
                </a:graphic>
              </wp:inline>
            </w:drawing>
          </mc:Choice>
          <mc:Fallback>
            <w:pict/>
          </mc:Fallback>
        </mc:AlternateContent>
      </w:r>
    </w:p>
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`; // 带文本框的圆角矩形绘图

    zip.file("[Content_Types].xml", contentTypesXml);
    zip.file("_rels/.rels", rootRelsXml);
    zip.file("word/document.xml", documentXml);

    return await zip.generateAsync({ type: "blob" });
  };

  it("should render wordprocessing shape drawing and textbox content", async () => {
    const docBlob = await createDrawingShapeDocx(); // 绘图测试文档
    const div = document.createElement("div"); // 渲染容器

    document.body.appendChild(div);
    await docx.renderAsync(docBlob, div);

    const shapeElement = div.querySelector("svg rect, svg ellipse, svg polygon, svg line"); // 任意已渲染形状节点
    const textContent = div.textContent || ""; // 页面文本内容

    expect(shapeElement).not.toBeNull();
    expect(textContent).toContain("绘图文本");

    div.remove();
  });
});
