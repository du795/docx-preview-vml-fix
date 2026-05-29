describe("Render VML stamp", function () {
  // 等待浏览器完成布局计算，便于读取 VML 的最终页面位置
  const waitForLayout = async () => {
    await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve))); // 等待双帧刷新完成布局
  };

  // 构造最小 VML 印章文档，覆盖 w:pict -> v:group -> v:textpath 场景
  const createVmlStampDocx = async () => {
    const zip = new JSZip(); // DOCX 压缩包对象
    const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`; // 内容类型清单
    const rootRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`; // 顶层关系定义
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:pict>
          <v:group id="_x0000_s2074" o:spid="_x0000_s2074" o:spt="203" style="position:absolute;left:0pt;margin-left:171.35pt;margin-top:9.55pt;height:82.05pt;width:135pt;" coordorigin="2095,8193" coordsize="2700,1641">
            <v:shape id="_x0000_s2075" o:spid="_x0000_s2075" o:spt="3" type="#_x0000_t3" style="position:absolute;left:2095;top:8193;height:1641;width:2700;" filled="f" stroked="t" coordsize="21600,21600">
              <v:fill on="f" focussize="0,0"/>
              <v:stroke weight="1.5pt" color="#FF0000"/>
            </v:shape>
            <v:group id="_x0000_s2076" o:spid="_x0000_s2076" o:spt="203" style="position:absolute;left:2380;top:8430;height:1280;width:2133;" coordorigin="2380,8430" coordsize="2133,1280">
              <v:shape id="_x0000_s2077" o:spid="_x0000_s2077" o:spt="144" type="#_x0000_t144" style="position:absolute;left:2380;top:8430;height:1022;width:2133;" filled="f" stroked="t" coordsize="21600,21600">
                <v:fill on="f" focussize="0,0"/>
                <v:stroke color="#FF0000"/>
                <v:textpath on="t" fitshape="t" fitpath="t" trim="t" xscale="f" string="国网重庆市北供电公司" style="font-family:宋体;font-size:12pt;v-text-align:center;"/>
              </v:shape>
              <v:shape id="_x0000_s2078" o:spid="_x0000_s2078" o:spt="136" type="#_x0000_t136" style="position:absolute;left:2575;top:8922;height:342;width:1785;" filled="f" stroked="t" coordsize="21600,21600">
                <v:fill on="f" focussize="0,0"/>
                <v:stroke color="#FF0000"/>
                <v:textpath on="t" fitshape="t" fitpath="t" trim="t" xscale="f" string="继 电 保 护 定 值" style="font-family:宋体;font-size:12pt;v-text-align:center;"/>
              </v:shape>
              <v:shape id="_x0000_s2079" o:spid="_x0000_s2079" o:spt="136" type="#_x0000_t136" style="position:absolute;left:2920;top:9350;height:360;width:1080;" filled="f" stroked="t" coordsize="21600,21600">
                <v:fill on="f" focussize="0,0"/>
                <v:stroke color="#FF0000"/>
                <v:textpath on="t" fitshape="t" fitpath="t" trim="t" xscale="f" string="专用章" style="font-family:宋体;font-size:18pt;v-text-align:center;"/>
              </v:shape>
            </v:group>
          </v:group>
        </w:pict>
      </w:r>
      <w:r>
        <w:t>印章文字测试</w:t>
      </w:r>
    </w:p>
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`; // 与用户样例同结构的 VML 印章节点

    zip.file("[Content_Types].xml", contentTypesXml);
    zip.file("_rels/.rels", rootRelsXml);
    zip.file("word/document.xml", documentXml);

    return await zip.generateAsync({ type: "blob" });
  };

  // 构造顶层 v:shape 文档，覆盖 w:pict -> v:shape 直接渲染场景
  const createTopLevelVmlShapeDocx = async () => {
    const zip = new JSZip(); // DOCX 压缩包对象
    const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`; // 内容类型清单
    const rootRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`; // 顶层关系定义
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:pict>
          <v:shape id="_x0000_s3001" o:spid="_x0000_s3001" o:spt="3" type="#_x0000_t3" style="position:absolute;left:36pt;top:24pt;height:72pt;width:72pt;" filled="f" stroked="t">
            <v:fill on="f" focussize="0,0"/>
            <v:stroke weight="1pt" color="#FF0000"/>
          </v:shape>
        </w:pict>
      </w:r>
    </w:p>
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`; // 顶层 VML 图形测试文档

    zip.file("[Content_Types].xml", contentTypesXml);
    zip.file("_rels/.rels", rootRelsXml);
    zip.file("word/document.xml", documentXml);

    return await zip.generateAsync({ type: "blob" });
  };

  it("should render vml stamp ellipse and text", async () => {
    const docBlob = await createVmlStampDocx(); // VML 印章测试文档
    const div = document.createElement("div"); // 页面渲染容器

    document.body.appendChild(div);
    await docx.renderAsync(docBlob, div);
    await waitForLayout(); // 等待浏览器完成 VML 的位置布局

    const ellipse = div.querySelector("svg ellipse"); // 印章椭圆节点
    const section = div.querySelector("section.docx"); // 页面节点
    const article = div.querySelector("section.docx > article"); // 正文内容区节点
    const rootSvg = div.querySelector("span > svg"); // 顶层 VML SVG 节点
    const textContent = div.textContent || ""; // 渲染结果文字
    const expectedLeft = 171.35 * 96 / 72; // 源文档相对页面左边缘的目标偏移量
    const pagePaddingLeft = parseFloat(getComputedStyle(section).paddingLeft); // 页面正文区域左内边距
    const actualLeft = rootSvg.getBoundingClientRect().left - article.getBoundingClientRect().left; // 顶层 VML 相对正文内容区左边缘的实际偏移量

    expect(ellipse).not.toBeNull();
    expect(section).not.toBeNull();
    expect(article).not.toBeNull();
    expect(rootSvg).not.toBeNull();
    expect(textContent).toContain("专用章");
    expect(textContent).toContain("国网重庆市北供电公司");
    expect(pagePaddingLeft).toBeGreaterThan(0);
    expect(Math.abs(actualLeft - expectedLeft)).toBeLessThan(2);

    div.remove();
  });

  it("should not double count top level vml shape left and top", async () => {
    const docBlob = await createTopLevelVmlShapeDocx(); // 顶层 VML 图形测试文档
    const div = document.createElement("div"); // 页面渲染容器

    document.body.appendChild(div);
    await docx.renderAsync(docBlob, div);

    const ellipse = div.querySelector("svg ellipse"); // 顶层椭圆节点
    const centerX = ellipse.getAttribute("cx"); // 椭圆中心横坐标
    const centerY = ellipse.getAttribute("cy"); // 椭圆中心纵坐标

    expect(ellipse).not.toBeNull();
    expect(centerX).toBe("36");
    expect(centerY).toBe("36");

    div.remove();
  });
});
