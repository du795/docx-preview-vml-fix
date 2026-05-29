import xml, { XmlParser } from "../parser/xml-parser";

export interface ChartSeriesData {
    name?: string;
    categories?: string[];
    values?: number[];
    xValues?: number[];
    color?: string;
    strokeColor?: string;
    pointColors?: string[];
}

export interface ChartData {
    type: string;
    plotTypes: string[];
    title?: string;
    categories?: string[];
    series: ChartSeriesData[];
    legendPosition?: string;
    categoryAxisTitle?: string;
    valueAxisTitle?: string;
    barDirection?: "bar" | "col";
    grouping?: string;
    holeSize?: number;
}

const knownChartTypes = new Set([
    "areaChart",
    "barChart",
    "bubbleChart",
    "doughnutChart",
    "lineChart",
    "ofPieChart",
    "pieChart",
    "radarChart",
    "scatterChart",
    "stockChart",
    "surfaceChart",
]);

// 解析图表部件根节点，抽取标题、图例、坐标轴标题、系列与缓存数据
export function parseChartSpace(root: Element, parser: XmlParser = xml): ChartData {
    var chartNode = parser.element(root, "chart"); // 图表主节点
    var plotAreaNode = chartNode ? parser.element(chartNode, "plotArea") : null; // 绘图区节点
    var plotNodes = plotAreaNode ? parser.elements(plotAreaNode).filter(isChartPlotNode) : []; // 图表类型节点列表
    var primaryPlotNode = plotNodes[0] ?? null; // 主图表节点
    var primaryType = primaryPlotNode?.localName ?? "unknownChart"; // 主图表类型
    var series = primaryPlotNode ? parseChartSeriesList(primaryPlotNode, primaryType, parser) : []; // 图表系列集合
    var categories = series.find(x => x.categories?.length > 0)?.categories ?? []; // 通用分类标签
    var legendNode = chartNode ? parser.element(chartNode, "legend") : null; // 图例节点
    var legendPositionNode = legendNode ? parser.element(legendNode, "legendPos") : null; // 图例位置节点
    var titleNode = chartNode ? parser.element(chartNode, "title") : null; // 图表标题节点
    var groupingNode = primaryPlotNode ? parser.element(primaryPlotNode, "grouping") : null; // 分组节点
    var holeSizeNode = primaryPlotNode ? parser.element(primaryPlotNode, "holeSize") : null; // 圆环图内孔节点

    return {
        type: primaryType,
        plotTypes: plotNodes.map(x => x.localName),
        title: parseChartText(titleNode, parser),
        categories,
        series,
        legendPosition: legendPositionNode ? parser.attr(legendPositionNode, "val") ?? "r" : (legendNode ? "r" : null),
        categoryAxisTitle: parseAxisTitle(plotAreaNode, "catAx", parser),
        valueAxisTitle: parseAxisTitle(plotAreaNode, "valAx", parser),
        barDirection: parseBarDirection(primaryPlotNode, parser),
        grouping: groupingNode ? parser.attr(groupingNode, "val") ?? "standard" : "standard",
        holeSize: holeSizeNode ? parser.intAttr(holeSizeNode, "val") : null,
    };
}

function isChartPlotNode(node: Element): boolean {
    return knownChartTypes.has(node.localName);
}

function parseBarDirection(node: Element, parser: XmlParser): "bar" | "col" {
    if (!node)
        return "col";

    var barDirectionNode = parser.element(node, "barDir"); // 柱状图方向节点
    return barDirectionNode ? parser.attr(barDirectionNode, "val") as any ?? "col" : "col";
}

// 解析单个图表节点下的系列集合，兼容分类图与散点图缓存结构
function parseChartSeriesList(node: Element, chartType: string, parser: XmlParser): ChartSeriesData[] {
    return parser.elements(node, "ser").map(serNode => {
        var shapePropertiesNode = parser.element(serNode, "spPr"); // 系列形状属性节点
        var categoriesNode = parser.element(serNode, "cat"); // 分类轴缓存节点
        var valuesNode = parser.element(serNode, "val") ?? parser.element(serNode, "yVal"); // 数值轴缓存节点
        var xValuesNode = parser.element(serNode, "xVal"); // 散点图横轴缓存节点

        return {
            name: parseSeriesName(serNode, parser),
            categories: chartType === "scatterChart" ? [] : parseStringDataValues(categoriesNode, parser),
            values: parseNumericDataValues(valuesNode, parser),
            xValues: chartType === "scatterChart" ? parseNumericDataValues(xValuesNode, parser) : [],
            color: parseSolidFillColor(shapePropertiesNode, parser),
            strokeColor: parseLineColor(shapePropertiesNode, parser),
            pointColors: parsePointColors(serNode, parser),
        };
    });
}

function parseSeriesName(node: Element, parser: XmlParser): string {
    var textNode = parser.element(node, "tx"); // 系列名称节点
    return parseChartText(textNode, parser);
}

// 解析图表文本，兼容 rich text、字符串缓存与直接值节点
function parseChartText(node: Element, parser: XmlParser): string {
    if (!node)
        return null;

    var directText = collectTextByLocalNames(node, ["t"], false); // 富文本中的文本片段
    var cacheText = collectTextByLocalNames(node, ["v"], true); // 缓存值文本

    return directText || cacheText || null;
}

function parseAxisTitle(plotAreaNode: Element, axisName: string, parser: XmlParser): string {
    var axisNode = plotAreaNode ? parser.elements(plotAreaNode, axisName)[0] : null; // 指定坐标轴节点
    return axisNode ? parseChartText(parser.element(axisNode, "title"), parser) : null;
}

// 解析分类文本缓存，兼容多级分类与字符串/数字缓存
function parseStringDataValues(node: Element, parser: XmlParser): string[] {
    if (!node)
        return [];

    var multiLevelNode = parser.element(node, "multiLvlStrRef"); // 多级分类引用节点

    if (multiLevelNode)
        return parseMultiLevelStringCache(parser.element(multiLevelNode, "multiLvlStrCache"), parser);

    var sourceNode = parser.element(node, "strRef")
        ?? parser.element(node, "strLit")
        ?? parser.element(node, "numRef")
        ?? parser.element(node, "numLit"); // 字符串或数字缓存来源节点
    var cacheNode = sourceNode ? parser.element(sourceNode, "strCache") ?? parser.element(sourceNode, "numCache") ?? sourceNode : null; // 实际缓存节点

    return parseIndexedTextCache(cacheNode, parser);
}

function parseNumericDataValues(node: Element, parser: XmlParser): number[] {
    if (!node)
        return [];

    var sourceNode = parser.element(node, "numRef")
        ?? parser.element(node, "numLit")
        ?? parser.element(node, "strRef")
        ?? parser.element(node, "strLit"); // 数值缓存来源节点
    var cacheNode = sourceNode ? parser.element(sourceNode, "numCache") ?? parser.element(sourceNode, "strCache") ?? sourceNode : null; // 实际缓存节点

    return parseIndexedTextCache(cacheNode, parser)
        .map(value => Number(value))
        .filter(value => !Number.isNaN(value));
}

function parseIndexedTextCache(node: Element, parser: XmlParser): string[] {
    if (!node)
        return [];

    var points = parser.elements(node, "pt").map(pointNode => ({ // 缓存点列表
        index: parser.intAttr(pointNode, "idx", 0),
        value: collectTextByLocalNames(pointNode, ["v"], true),
    }));

    return points
        .sort((left, right) => left.index - right.index)
        .map(point => point.value ?? "");
}

function parseMultiLevelStringCache(node: Element, parser: XmlParser): string[] {
    if (!node)
        return [];

    var levelMaps = parser.elements(node, "lvl").map(levelNode => parseIndexedTextCache(levelNode, parser)); // 多层级文本缓存
    var maxLength = levelMaps.reduce((maxValue, currentValue) => Math.max(maxValue, currentValue.length), 0); // 最大分类数量
    var result = new Array(maxLength).fill("").map((_, index) => levelMaps
        .map(level => level[index])
        .filter(Boolean)
        .join(" / ")); // 多级分类合并文本

    return result.filter(Boolean);
}

function parsePointColors(node: Element, parser: XmlParser): string[] {
    var result: string[] = []; // 饼图等按点着色集合

    parser.elements(node, "dPt").forEach(pointNode => {
        var pointIndexNode = parser.element(pointNode, "idx"); // 当前数据点索引节点
        var pointIndex = pointIndexNode ? parser.intAttr(pointIndexNode, "val", -1) : -1; // 当前数据点索引
        var pointColor = parseSolidFillColor(parser.element(pointNode, "spPr"), parser); // 当前数据点填充色

        if (pointIndex >= 0 && pointColor)
            result[pointIndex] = pointColor;
    });

    return result;
}

function parseSolidFillColor(node: Element, parser: XmlParser): string {
    var solidFillNode = node ? parser.element(node, "solidFill") : null; // 填充颜色节点
    return parseColorValue(solidFillNode, parser);
}

function parseLineColor(node: Element, parser: XmlParser): string {
    var lineNode = node ? parser.element(node, "ln") : null; // 线条节点
    var solidFillNode = lineNode ? parser.element(lineNode, "solidFill") : null; // 线条填充节点
    return parseColorValue(solidFillNode, parser);
}

function parseColorValue(node: Element, parser: XmlParser): string {
    if (!node)
        return null;

    var rgbNode = parser.element(node, "srgbClr"); // 直接 RGB 颜色节点
    var schemeNode = parser.element(node, "schemeClr"); // 主题颜色节点
    var presetNode = parser.element(node, "prstClr"); // 预设颜色节点

    if (rgbNode)
        return `#${parser.attr(rgbNode, "val")}`;

    if (schemeNode)
        return `var(--docx-${parser.attr(schemeNode, "val")}-color)`;

    if (presetNode)
        return parser.attr(presetNode, "val");

    return null;
}

function collectTextByLocalNames(node: Element, localNames: string[], useLineBreak: boolean): string {
    var fragments: string[] = []; // 文本片段集合

    walkElements(node, currentNode => {
        if (!localNames.includes(currentNode.localName))
            return;

        var currentText = currentNode.textContent ?? ""; // 当前节点文本

        if (currentText)
            fragments.push(currentText);
    });

    return fragments.join(useLineBreak ? "" : "").trim();
}

function walkElements(node: Element, visitor: (node: Element) => void) {
    visitor(node);

    for (let childNode of Array.from(node.children)) {
        walkElements(childNode, visitor);
    }
}
