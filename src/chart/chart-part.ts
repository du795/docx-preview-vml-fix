import { Part } from "../common/part";
import { OpenXmlPackage } from "../common/open-xml-package";
import { ChartData, parseChartSpace } from "./chart";

export class ChartPart extends Part {
    chart: ChartData;

    constructor(pkg: OpenXmlPackage, path: string) {
        super(pkg, path);
    }

    // 解析图表部件 XML，并缓存为后续渲染使用的图表数据对象
    parseXml(root: Element) {
        this.chart = parseChartSpace(root);
    }
}
