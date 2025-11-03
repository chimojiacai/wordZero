// @Author: liyongzhen
// @Description:
// @File: section
// @Date: 2025/6/25 17:48
// 📁 文件: pkg/document/section.go

package document

import (
	"encoding/xml"
	"fmt"
)

// AddSectionBreak 用于生成 orientation
func (p *Paragraph) AddSectionBreak(orient PageOrientation, doc *Document) {
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
		// 继承现有的页眉页脚引用
		HeaderReferences: existingSectPr.HeaderReferences,
		FooterReferences: existingSectPr.FooterReferences,
		TitlePage:        existingSectPr.TitlePage,
		PageNumType:      existingSectPr.PageNumType,
		XmlnsR:           existingSectPr.XmlnsR,
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
