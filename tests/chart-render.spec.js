describe("Render chart", function () {
  // 构造最小图表 DOCX 文档，覆盖 drawing -> graphicData -> c:chart 关系加载与渲染链路
  const createChartDocx = async ({ chartXml, chartRelationshipId = "rIdChart1" }) => {
    const zip = new JSZip(); // DOCX 压缩包对象
    const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
</Types>`; // 内容类型清单
    const rootRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`; // 顶层关系定义
    const documentRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="${chartRelationshipId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/>
</Relationships>`; // 文档与图表关系
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <w:body>
    <w:p>
      <w:r>
        <w:drawing>
          <wp:inline>
            <wp:extent cx="4114800" cy="2743200"/>
            <wp:docPr id="1" name="Chart 1"/>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                <c:chart r:id="${chartRelationshipId}"/>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`; // 内嵌图表文档结构

    zip.file("[Content_Types].xml", contentTypesXml);
    zip.file("_rels/.rels", rootRelsXml);
    zip.file("word/document.xml", documentXml);
    zip.file("word/_rels/document.xml.rels", documentRelsXml);
    zip.file("word/charts/chart1.xml", chartXml);

    return await zip.generateAsync({ type: "blob" });
  };

  it("should render bar chart as svg", async () => {
    const barChartXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:title>
      <c:tx>
        <c:rich>
          <a:p><a:r><a:t>季度销售额</a:t></a:r></a:p>
        </c:rich>
      </c:tx>
    </c:title>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx><c:v>系列A</c:v></c:tx>
          <c:spPr><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></c:spPr>
          <c:cat>
            <c:strLit>
              <c:ptCount val="3"/>
              <c:pt idx="0"><c:v>一月</c:v></c:pt>
              <c:pt idx="1"><c:v>二月</c:v></c:pt>
              <c:pt idx="2"><c:v>三月</c:v></c:pt>
            </c:strLit>
          </c:cat>
          <c:val>
            <c:numLit>
              <c:ptCount val="3"/>
              <c:pt idx="0"><c:v>12</c:v></c:pt>
              <c:pt idx="1"><c:v>18</c:v></c:pt>
              <c:pt idx="2"><c:v>10</c:v></c:pt>
            </c:numLit>
          </c:val>
        </c:ser>
        <c:axId val="1"/>
        <c:axId val="2"/>
      </c:barChart>
      <c:catAx>
        <c:axId val="1"/>
        <c:title><c:tx><c:rich><a:p><a:r><a:t>月份</a:t></a:r></a:p></c:rich></c:tx></c:title>
      </c:catAx>
      <c:valAx>
        <c:axId val="2"/>
        <c:title><c:tx><c:rich><a:p><a:r><a:t>金额</a:t></a:r></a:p></c:rich></c:tx></c:title>
      </c:valAx>
    </c:plotArea>
    <c:legend><c:legendPos val="r"/></c:legend>
  </c:chart>
</c:chartSpace>`; // 柱状图 XML
    const docBlob = await createChartDocx({ chartXml: barChartXml }); // 图表测试文档
    const div = document.createElement("div"); // 页面渲染容器

    document.body.appendChild(div);
    await docx.renderAsync(docBlob, div);

    const chartNode = div.querySelector(".docx-chart"); // 图表容器节点
    const rectNode = div.querySelector(".docx-chart svg rect"); // 柱状图矩形节点
    const textContent = div.textContent || ""; // 页面文本内容

    expect(chartNode).not.toBeNull();
    expect(rectNode).not.toBeNull();
    expect(textContent).toContain("季度销售额");
    expect(textContent).toContain("系列A");
    expect(textContent).toContain("一月");

    div.remove();
  });

  it("should fallback unsupported chart to data table", async () => {
    const radarChartXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:title>
      <c:tx>
        <c:rich>
          <a:p><a:r><a:t>雷达图测试</a:t></a:r></a:p>
        </c:rich>
      </c:tx>
    </c:title>
    <c:plotArea>
      <c:radarChart>
        <c:radarStyle val="marker"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx><c:v>能力值</c:v></c:tx>
          <c:cat>
            <c:strLit>
              <c:ptCount val="3"/>
              <c:pt idx="0"><c:v>沟通</c:v></c:pt>
              <c:pt idx="1"><c:v>执行</c:v></c:pt>
              <c:pt idx="2"><c:v>质量</c:v></c:pt>
            </c:strLit>
          </c:cat>
          <c:val>
            <c:numLit>
              <c:ptCount val="3"/>
              <c:pt idx="0"><c:v>80</c:v></c:pt>
              <c:pt idx="1"><c:v>90</c:v></c:pt>
              <c:pt idx="2"><c:v>85</c:v></c:pt>
            </c:numLit>
          </c:val>
        </c:ser>
      </c:radarChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`; // 暂不支持的雷达图 XML
    const docBlob = await createChartDocx({ chartXml: radarChartXml }); // 图表测试文档
    const div = document.createElement("div"); // 页面渲染容器

    document.body.appendChild(div);
    await docx.renderAsync(docBlob, div);

    const tableNode = div.querySelector(".docx-chart table"); // 数据回退表格
    const textContent = div.textContent || ""; // 页面文本内容

    expect(tableNode).not.toBeNull();
    expect(textContent).toContain("雷达图测试");
    expect(textContent).toContain("能力值");
    expect(textContent).toContain("暂未完整支持 radarChart");

    div.remove();
  });
});
