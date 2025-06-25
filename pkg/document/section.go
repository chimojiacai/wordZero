// @Author: liyongzhen
// @Description:
// @File: section
// @Date: 2025/6/25 17:48
// 📁 文件: pkg/document/section.go

package document

import "encoding/xml"

// AddSectionBreak 用于生成 orientation
func (p *Paragraph) AddSectionBreak(orient PageOrientation) {
	if p.Properties == nil {
		p.Properties = &ParagraphProperties{}
	}

	sectPr := &SectionProperties{
		XMLName:  xml.Name{Local: "w:sectPr"},
		PageSize: &PageSizeXML{},
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
