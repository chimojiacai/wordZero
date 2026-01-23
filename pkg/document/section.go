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

	// Word分节符的逻辑：分节符存储的是它*之前*那一部分内容的节属性。
	// 当我们添加一个分节符时，我们实际上是把“当前节”的属性打包存入分节符中，
	// 然后文档后续部分开始一个新的节（新节属性存储在文档末尾的sectPr或者下一个分节符中）。
	//
	// 所以，这里的 sectPr 实际上是用来定义 *分节符之前* 的那一部分内容的格式。
	//
	// 但是，startPage (页码起始值) 这个属性比较特殊。
	// 在Word XML中，sectPr.PageNumType.Start 定义的是 *本节* 的起始页码。
	//
	// 如果我们希望 *分节符之后* 的新节从第1页开始：
	// 我们需要设置的是 *新节* 的属性，而不是分节符里的属性（分节符里的属性控制的是上一节）。
	//
	// 然而，sectPr 是当前段落的属性。在Word中，段落末尾的分节符确实定义了 *本节*（即分节符所在节，也就是上一节）的属性。
	// 下一节的属性由下一节结尾的分节符或文档末尾的sectPr定义。
	//
	// 现在的需求是：在 sectionBreak1 之后的内容（即下一节）页码从1开始。
	//
	// 在 TestProductionIssue 中：
	// sectionBreak1 := doc.AddParagraph("")
	// sectionBreak1.AddSectionBreakWithStartPage(..., startPage=1, ...)
	//
	// 这段代码创建了一个分节符。如果我们将 Start=1 放在这个分节符的 sectPr 中，
	// 那么它将应用于 *sectionBreak1 之前* 的内容（即封面、目录等）。
	// 这显然不是我们想要的。我们想要的是 *sectionBreak1 之后* 的内容从1开始。
	//
	// 因此，如果我们要设置下一节的起始页码，我们需要修改的是 *下一节* 的 sectPr。
	// 在我们的实现中，下一节的 sectPr 位于文档末尾（doc.Body.Elements的最后，或者是动态获取的）。
	//
	// 所以，这个函数的逻辑需要调整：
	// 1. 创建分节符（sectPr），用于结束当前节（上一节）。
	//    对于上一节，我们通常保持默认（不重置页码），或者继承之前的设置。
	// 2. 准备下一节的属性。如果 startPage > 0，这意味着 *下一节* 需要重置页码。
	//    我们需要找到或创建下一节的 sectPr（文档末尾的），并设置它的 PageNumType.Start。
	
	// 1. 创建当前节（上一节）的属性，用于分节符
	prevSectPr := &SectionProperties{
		XMLName:  xml.Name{Local: "w:sectPr"},
		PageSize: &PageSizeXML{},
		PageMargins: &PageMargin{
			XMLName: xml.Name{Local: "w:pgMar"},
			Top:     fmt.Sprintf("%.0f", mmToTwips(doc.GetPageSettings().MarginTop)),
			Bottom:  fmt.Sprintf("%.0f", mmToTwips(doc.GetPageSettings().MarginBottom)),
			Left:    fmt.Sprintf("%.0f", mmToTwips(doc.GetPageSettings().MarginLeft)),
			Right:   fmt.Sprintf("%.0f", mmToTwips(doc.GetPageSettings().MarginRight)),
			Header:  fmt.Sprintf("%.0f", mmToTwips(doc.GetPageSettings().HeaderDistance)),
			Footer:  fmt.Sprintf("%.0f", mmToTwips(doc.GetPageSettings().FooterDistance)),
		},
		TitlePage: existingSectPr.TitlePage,
		XmlnsR:    existingSectPr.XmlnsR,
		PageNumType: existingSectPr.PageNumType, // 继承上一节的页码设置（通常是默认连续）
		Columns: existingSectPr.Columns,
		DocGrid: existingSectPr.DocGrid,
	}

	// 继承页眉页脚引用（用于上一节）
	// 注意：这里的 inheritHeaderFooter 参数通常是用户想要对 *新节* 做的设置。
	// 但在这里，我们正在构建的是 *上一节* 的结束符。
	// 上一节的页眉页脚应该保持原样（即 existingSectPr 中的内容）。
	if existingSectPr.HeaderReferences != nil {
		prevSectPr.HeaderReferences = make([]*HeaderFooterReference, len(existingSectPr.HeaderReferences))
		copy(prevSectPr.HeaderReferences, existingSectPr.HeaderReferences)
	}
	if existingSectPr.FooterReferences != nil {
		prevSectPr.FooterReferences = make([]*FooterReference, len(existingSectPr.FooterReferences))
		copy(prevSectPr.FooterReferences, existingSectPr.FooterReferences)
	}
	
	// 设置上一节的页面方向
	// 注意：orient 参数通常也是用户想要设置的 *新节* 方向。
	// 上一节的方向应该保持原样（doc.GetPageSettings().Orientation）。
	// 但为了简单起见，目前的实现似乎是假设 AddSectionBreak 是为了改变后续方向，
	// 而 Word 的行为是分节符定义了它 *所在节* 的属性。
	// 如果我们想改变下一节的方向，我们需要修改下一节的 sectPr。
	//
	// 让我们假设 existingSectPr 包含了当前节（即将结束的节）的正确属性。
	// 我们直接复制它作为分节符属性。
	if existingSectPr.PageSize != nil {
		prevSectPr.PageSize = &PageSizeXML{
			W: existingSectPr.PageSize.W,
			H: existingSectPr.PageSize.H,
			Orient: existingSectPr.PageSize.Orient,
		}
	} else {
		// 默认 A4 Portrait
		prevSectPr.PageSize = &PageSizeXML{
			W: "11906",
			H: "16838",
			Orient: "portrait",
		}
	}

	// 将 prevSectPr 赋值给分节符段落
	p.Properties.SectionProperties = prevSectPr


	// 2. 准备下一节（新节）的属性
	// 下一节的属性存储在文档末尾的 sectPr 中。
	// 我们需要获取它（如果不存在则创建），并应用新的设置（startPage, orient, inheritHeaderFooter）。
	
	// 获取文档末尾的 sectPr（这控制新的一节）
	// 注意：getCurrentSectionProperties 现在优先返回文档末尾的 sectPr。
	nextSectPr := doc.getCurrentSectionProperties()
	
	// 更新新节的页面方向
	if nextSectPr.PageSize == nil {
		nextSectPr.PageSize = &PageSizeXML{}
	}
	if orient == OrientationLandscape {
		nextSectPr.PageSize.Orient = "landscape"
		nextSectPr.PageSize.W = "16838" // landscape A4
		nextSectPr.PageSize.H = "11906"
	} else {
		nextSectPr.PageSize.Orient = "portrait"
		nextSectPr.PageSize.W = "11906"
		nextSectPr.PageSize.H = "16838"
	}

	// 更新新节的页面边距（包括页眉页脚距离）
	if nextSectPr.PageMargins == nil {
		nextSectPr.PageMargins = &PageMargin{}
	}
	nextSectPr.PageMargins.Top = fmt.Sprintf("%.0f", mmToTwips(doc.GetPageSettings().MarginTop))
	nextSectPr.PageMargins.Bottom = fmt.Sprintf("%.0f", mmToTwips(doc.GetPageSettings().MarginBottom))
	nextSectPr.PageMargins.Left = fmt.Sprintf("%.0f", mmToTwips(doc.GetPageSettings().MarginLeft))
	nextSectPr.PageMargins.Right = fmt.Sprintf("%.0f", mmToTwips(doc.GetPageSettings().MarginRight))
	nextSectPr.PageMargins.Header = fmt.Sprintf("%.0f", mmToTwips(doc.GetPageSettings().HeaderDistance))
	nextSectPr.PageMargins.Footer = fmt.Sprintf("%.0f", mmToTwips(doc.GetPageSettings().FooterDistance))

	// 更新新节的页码设置
	nextSectPr.PageNumType = &PageNumType{
		Fmt: "decimal",
	}
	if startPage > 0 {
		nextSectPr.PageNumType.Start = strconv.Itoa(startPage)
	} else {
		// 如果 startPage == 0，表示我们想要保持页码连续。
		// Word 中，如果 <w:pgNumType> 不包含 w:start 属性，默认就是连续的 (continue from previous section)。
		// 但是，如果我们是从一个设置了 start 的节跳转到一个没有设置 start 的节，
		// 某些版本的 Word 可能会有问题。
		// 无论如何，如果我们显式创建了 PageNumType，我们应该确保它是正确的。
		//
		// 关键点：如果是为了延续页码，我们不应该设置 Start 属性。
		// nextSectPr.PageNumType.Start 应该为空。
		// 现有的逻辑 (if startPage > 0) 已经保证了这一点。
		//
		// 但是，如果 nextSectPr 之前已经有了 Start 属性（比如被复制过来的），我们需要清除它。
		nextSectPr.PageNumType.Start = ""
	}

	// 处理页眉页脚继承
	if !inheritHeaderFooter {
		// 如果不继承，清除新节的页眉页脚引用
		nextSectPr.HeaderReferences = nil
		nextSectPr.FooterReferences = nil
	} else {
		// 如果继承，确保新节拥有与上一节相同的引用
		// (如果 nextSectPr 是新建的，它可能为空；如果是获取到的 existing，它可能已经有了)
		if len(nextSectPr.HeaderReferences) == 0 && len(existingSectPr.HeaderReferences) > 0 {
			nextSectPr.HeaderReferences = make([]*HeaderFooterReference, len(existingSectPr.HeaderReferences))
			copy(nextSectPr.HeaderReferences, existingSectPr.HeaderReferences)
		}
		if len(nextSectPr.FooterReferences) == 0 && len(existingSectPr.FooterReferences) > 0 {
			nextSectPr.FooterReferences = make([]*FooterReference, len(existingSectPr.FooterReferences))
			copy(nextSectPr.FooterReferences, existingSectPr.FooterReferences)
		}
	}
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
