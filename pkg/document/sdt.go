// Package document 提供Word文档的SDT（Structured Document Tag）结构
package document

import (
	"encoding/xml"
	"fmt"
)

// SDT 结构化文档标签，用于目录等特殊功能
type SDT struct {
	XMLName    xml.Name       `xml:"w:sdt"`
	Properties *SDTProperties `xml:"w:sdtPr"`
	EndPr      *SDTEndPr      `xml:"w:sdtEndPr,omitempty"`
	Content    *SDTContent    `xml:"w:sdtContent"`
}

// SDTProperties SDT属性
type SDTProperties struct {
	XMLName     xml.Name        `xml:"w:sdtPr"`
	RunPr       *RunProperties  `xml:"w:rPr,omitempty"`
	ID          *SDTID          `xml:"w:id,omitempty"`
	Color       *SDTColor       `xml:"w15:color,omitempty"`
	DocPartObj  *DocPartObj     `xml:"w:docPartObj,omitempty"`
	Placeholder *SDTPlaceholder `xml:"w:placeholder,omitempty"`
}

// SDTEndPr SDT结束属性
type SDTEndPr struct {
	XMLName xml.Name       `xml:"w:sdtEndPr"`
	RunPr   *RunProperties `xml:"w:rPr,omitempty"`
}

// SDTContent SDT内容
type SDTContent struct {
	XMLName  xml.Name      `xml:"w:sdtContent"`
	Elements []interface{} `xml:"-"` // 使用自定义序列化
}

// MarshalXML 自定义XML序列化
func (s *SDTContent) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	// 开始元素
	if err := e.EncodeToken(start); err != nil {
		return err
	}

	// 序列化每个元素
	for _, element := range s.Elements {
		if err := e.Encode(element); err != nil {
			return err
		}
	}

	// 结束元素
	return e.EncodeToken(start.End())
}

// SDTID SDT标识符
type SDTID struct {
	XMLName xml.Name `xml:"w:id"`
	Val     string   `xml:"w:val,attr"`
}

// SDTColor SDT颜色
type SDTColor struct {
	XMLName xml.Name `xml:"w15:color"`
	Val     string   `xml:"w:val,attr"`
}

// DocPartObj 文档部件对象
type DocPartObj struct {
	XMLName        xml.Name        `xml:"w:docPartObj"`
	DocPartGallery *DocPartGallery `xml:"w:docPartGallery,omitempty"`
	DocPartUnique  *DocPartUnique  `xml:"w:docPartUnique,omitempty"`
}

// DocPartGallery 文档部件库
type DocPartGallery struct {
	XMLName xml.Name `xml:"w:docPartGallery"`
	Val     string   `xml:"w:val,attr"`
}

// DocPartUnique 文档部件唯一标识
type DocPartUnique struct {
	XMLName xml.Name `xml:"w:docPartUnique"`
}

// SDTPlaceholder SDT占位符
type SDTPlaceholder struct {
	XMLName xml.Name `xml:"w:placeholder"`
	DocPart *DocPart `xml:"w:docPart,omitempty"`
}

// DocPart 文档部件
type DocPart struct {
	XMLName xml.Name `xml:"w:docPart"`
	Val     string   `xml:"w:val,attr"`
}

// Tab 制表符
type Tab struct {
	XMLName xml.Name `xml:"w:tab"`
}

// 实现BodyElement接口
func (s *SDT) ElementType() string {
	return "sdt"
}

// CreateTOCSDT 创建目录SDT结构
func (d *Document) CreateTOCSDT(title string, maxLevel int) *SDT {
	sdt := &SDT{
		Properties: &SDTProperties{
			RunPr: &RunProperties{
				FontFamily: &FontFamily{ASCII: "宋体"},
				FontSize:   &FontSize{Val: "21"},
			},
			ID:    &SDTID{Val: "147476628"},
			Color: &SDTColor{Val: "DBDBDB"},
			DocPartObj: &DocPartObj{
				DocPartGallery: &DocPartGallery{Val: "Table of Contents"},
				DocPartUnique:  &DocPartUnique{},
			},
		},
		EndPr: &SDTEndPr{
			RunPr: &RunProperties{
				FontSize: &FontSize{Val: "20"},
			},
		},
		Content: &SDTContent{
			Elements: []interface{}{},
		},
	}

	// 添加目录标题段落
	titlePara := &Paragraph{
		Properties: &ParagraphProperties{
			Spacing: &Spacing{
				Before: "0",
				After:  "0",
				Line:   "240",
			},
			Indentation: &Indentation{
				Left:      "0",
				Right:     "0",
				FirstLine: "0",
			},
			Justification: &Justification{Val: "center"},
		},
		Runs: []Run{
			{
				Text: Text{Content: title},
				Properties: &RunProperties{
					FontFamily: &FontFamily{ASCII: "宋体"},
					FontSize:   &FontSize{Val: "21"},
				},
			},
		},
	}

	// 添加书签开始 - 使用已有的BookmarkStart类型
	bookmarkStart := &BookmarkStart{
		ID:   "0",
		Name: "_Toc11693_WPSOffice_Type3",
	}

	sdt.Content.Elements = append(sdt.Content.Elements, bookmarkStart, titlePara)

	return sdt
}

// AddTOCEntry 向目录SDT添加条目（支持超链接跳转）
func (sdt *SDT) AddTOCEntry(text string, level int, pageNum int, bookmarkID string) {
	// 确定目录样式ID (13=toc 1, 14=toc 2, 15=toc 3等)
	styleVal := fmt.Sprintf("%d", 12+level)

	// 创建目录条目段落
	entryPara := &Paragraph{
		Properties: &ParagraphProperties{
			ParagraphStyle: &ParagraphStyle{Val: styleVal},
			Tabs: &Tabs{
				Tabs: []TabDef{
					{
						Val:    "right",
						Leader: "dot",
						Pos:    "8640",
					},
				},
			},
		},
		Runs: []Run{},
	}

	// 使用超链接域字段来支持跳转
	// 创建超链接域开始（不设置颜色，使用默认黑色）
	entryPara.Runs = append(entryPara.Runs, Run{
		FieldChar: &FieldChar{
			FieldCharType: "begin",
		},
	})

	// 添加超链接指令
	entryPara.Runs = append(entryPara.Runs, Run{
		InstrText: &InstrText{
			Space:   "preserve",
			Content: fmt.Sprintf(" HYPERLINK \\l \"%s\" ", bookmarkID),
		},
	})

	// 超链接域分隔符
	entryPara.Runs = append(entryPara.Runs, Run{
		FieldChar: &FieldChar{
			FieldCharType: "separate",
		},
	})

	// 添加标题文本（超链接显示文本，使用黑色）
	entryPara.Runs = append(entryPara.Runs, Run{
		Properties: &RunProperties{
			FontFamily: &FontFamily{ASCII: "Calibri", HAnsi: "Calibri", EastAsia: "宋体"},
			FontSize:   &FontSize{Val: "22"},
			Color:      &Color{Val: "000000"}, // 黑色
		},
		Text: Text{Content: text},
	})

	// 添加制表符
	entryPara.Runs = append(entryPara.Runs, Run{
		Text: Text{Content: "\t"},
	})

	// 添加页码引用域（使用PAGEREF域自动获取页码）
	entryPara.Runs = append(entryPara.Runs, Run{
		FieldChar: &FieldChar{
			FieldCharType: "begin",
		},
	})

	entryPara.Runs = append(entryPara.Runs, Run{
		InstrText: &InstrText{
			Space:   "preserve",
			Content: fmt.Sprintf(" PAGEREF %s \\h ", bookmarkID),
		},
	})

	entryPara.Runs = append(entryPara.Runs, Run{
		FieldChar: &FieldChar{
			FieldCharType: "separate",
		},
	})

	// 页码文本（初始值，Word会自动通过PAGEREF域更新为正确页码）
	// PAGEREF域会自动获取书签所在页的页码，所以初始值不重要
	entryPara.Runs = append(entryPara.Runs, Run{
		Properties: &RunProperties{
			FontFamily: &FontFamily{ASCII: "Calibri", HAnsi: "Calibri", EastAsia: "宋体"},
			FontSize:   &FontSize{Val: "22"},
		},
		Text: Text{Content: fmt.Sprintf("%d", pageNum)}, // 初始值，Word会自动更新为正确页码
	})

	// 页码域结束
	entryPara.Runs = append(entryPara.Runs, Run{
		FieldChar: &FieldChar{
			FieldCharType: "end",
		},
	})

	// 超链接域结束（不设置颜色，使用默认黑色）
	entryPara.Runs = append(entryPara.Runs, Run{
		FieldChar: &FieldChar{
			FieldCharType: "end",
		},
	})

	// 添加段落到SDT内容中
	sdt.Content.Elements = append(sdt.Content.Elements, entryPara)
}

// generatePlaceholderGUID 生成占位符GUID
func generatePlaceholderGUID(level int) string {
	guids := map[int]string{
		1: "{b5fdec38-8301-4b26-9716-d8b31c00c718}",
		2: "{a500490c-aaae-4252-8340-aa59729b9870}",
		3: "{d7310822-77d9-4e43-95e1-4649f1e215b3}",
	}

	if guid, exists := guids[level]; exists {
		return guid
	}
	return "{b5fdec38-8301-4b26-9716-d8b31c00c718}" // 默认使用1级
}

// FinalizeTOCSDT 完成目录SDT构建
func (sdt *SDT) FinalizeTOCSDT() {
	// 添加书签结束 - 使用已有的BookmarkEnd类型
	bookmarkEnd := &BookmarkEnd{
		ID: "0",
	}
	sdt.Content.Elements = append(sdt.Content.Elements, bookmarkEnd)
}
