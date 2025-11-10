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
// 创建新节时，默认会重置页码从1开始（除非明确设置了PageNumType.Start）
func (p *Paragraph) AddSectionBreak(orient PageOrientation, doc *Document) {
	if p.Properties == nil {
		p.Properties = &ParagraphProperties{}
	}

	// 在添加分节符之前，清除文档末尾节属性中的页眉页脚引用
	// 这样分节符之前的内容（封面、目录等）就不会显示页眉页脚
	doc.clearHeaderFooterReferences()

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
		// 不继承页眉页脚引用，创建新节时不包含页眉页脚
		// 这样分节符之前的内容不会显示页眉页脚
		// 只有在新节中调用AddStyleHeader和AddFooterWithPageNumber后才会显示
		HeaderReferences: nil,
		FooterReferences: nil,
		TitlePage:        existingSectPr.TitlePage,
		XmlnsR:           existingSectPr.XmlnsR,
	}

	// 创建新的PageNumType，默认从1开始
	// 这样新节的页码会从1开始，而不是继承旧节的页码
	sectPr.PageNumType = &PageNumType{
		Fmt:   "decimal",
		Start: "1", // 默认从1开始
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
func (p *Paragraph) AddSectionBreakWithPageNumber(orient PageOrientation, doc *Document, startPage int) {
	if p.Properties == nil {
		p.Properties = &ParagraphProperties{}
	}

	// 在添加分节符之前，清除文档末尾节属性中的页眉页脚引用
	// 这样分节符之前的内容（封面、目录等）就不会显示页眉页脚
	doc.clearHeaderFooterReferences()

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
		// 不继承页眉页脚引用，创建新节时不包含页眉页脚
		// 这样分节符之前的内容不会显示页眉页脚
		// 只有在新节中调用AddStyleHeader和AddFooterWithPageNumber后才会显示
		HeaderReferences: nil,
		FooterReferences: nil,
		TitlePage:        existingSectPr.TitlePage,
		XmlnsR:           existingSectPr.XmlnsR,
	}

	// 设置指定的起始页码
	sectPr.PageNumType = &PageNumType{
		Fmt:   "decimal",
		Start: strconv.Itoa(startPage),
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
