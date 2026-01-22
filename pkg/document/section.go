// @Author: liyongzhen
// @Description:
// @File: section
// @Date: 2025/6/25 17:48
// 📁 文件: pkg/document/section.go

package document

import (
	"encoding/xml"
	"fmt"
	"strconv"
)

// AddSectionBreak 用于生成 orientation
// 创建新节时，如果不希望重置页码，需要额外处理
func (p *Paragraph) AddSectionBreak(orient PageOrientation, doc *Document) {
	p.AddSectionBreakWithStartPage(orient, doc, 0, true)
}

// AddSectionBreakWithStartPage 添加分节符并指定起始页码
// 参数:
//   - orient: 页面方向
//   - doc: 文档对象
//   - startPage: 起始页码，0表示延续上一节的页码
//   - inheritHeaderFooter: 是否继承上一节的页眉页脚
func (p *Paragraph) AddSectionBreakWithStartPage(orient PageOrientation, doc *Document, startPage int, inheritHeaderFooter bool) {
	if p.Properties == nil {
		p.Properties = &ParagraphProperties{}
	}

	// 获取现有的节属性（如果有）
	existingSectPr := doc.getSectionPropertiesForHeaderFooter()

	sectPr := &SectionProperties{
		XMLName:  xml.Name{Local: "w:sectPr"},
		PageSize: &PageSizeXML{},
		PageMargins: &PageMargin{
			XMLName: xml.Name{Local: "w:pgMar"},
			Top:     fmt.Sprintf("%.0f", mmToTwips(doc.GetPageSettings().MarginTop)),
			Bottom:  fmt.Sprintf("%.0f", mmToTwips(doc.GetPageSettings().MarginBottom)),
			Left:    fmt.Sprintf("%.0f", mmToTwips(doc.GetPageSettings().MarginLeft)),
			Right:   fmt.Sprintf("%.0f", mmToTwips(doc.GetPageSettings().MarginRight)),
		},
		TitlePage: existingSectPr.TitlePage,
		XmlnsR:    existingSectPr.XmlnsR,
	}

	// 继承页眉页脚引用
	if inheritHeaderFooter {
		if existingSectPr.HeaderReferences != nil {
			sectPr.HeaderReferences = make([]*HeaderFooterReference, len(existingSectPr.HeaderReferences))
			copy(sectPr.HeaderReferences, existingSectPr.HeaderReferences)
		}
		if existingSectPr.FooterReferences != nil {
			sectPr.FooterReferences = make([]*FooterReference, len(existingSectPr.FooterReferences))
			copy(sectPr.FooterReferences, existingSectPr.FooterReferences)
		}
	}

	// 设置页码类型
	sectPr.PageNumType = &PageNumType{
		Fmt: "decimal",
	}
	if startPage > 0 {
		sectPr.PageNumType.Start = strconv.Itoa(startPage)
	}

	if orient == OrientationLandscape {
		sectPr.PageSize.Orient = "landscape"
		sectPr.PageSize.W = "16838" // landscape A4
		sectPr.PageSize.H = "11906"
	} else {
		sectPr.PageSize.Orient = "portrait"
		sectPr.PageSize.W = "11906"
		sectPr.PageSize.H = "16838"
	}

	p.Properties.SectionProperties = sectPr
}

// AddSectionBreakWithPageNumber 添加分节符并设置起始页码
// 注意：此方法已弃用，请使用 AddSectionBreakWithStartPage 替代
func (p *Paragraph) AddSectionBreakWithPageNumber(orient PageOrientation, doc *Document, startPage int) {
	p.AddSectionBreakWithStartPage(orient, doc, startPage, false)
}

// AddSectionBreakContinuous 添加分节符但保持页码连续
// 注意：此方法已弃用，请使用 AddSectionBreakWithStartPage 替代
func (p *Paragraph) AddSectionBreakContinuous(orient PageOrientation, doc *Document) {
	p.AddSectionBreakWithStartPage(orient, doc, 0, true)
}
