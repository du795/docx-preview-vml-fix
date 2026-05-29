import { OpenXmlElementBase, DomType } from "../document/dom";

export abstract class WmlBaseNote implements OpenXmlElementBase {
    abstract type: DomType; // 抽象批注类型由具体脚注或尾注子类赋值
    id: string;
	noteType: string;
}

export class WmlFootnote extends WmlBaseNote {
	type = DomType.Footnote
}

export class WmlEndnote extends WmlBaseNote {
	type = DomType.Endnote
}
