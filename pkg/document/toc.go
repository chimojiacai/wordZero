// Package document 提供Word文档目录生成功能
package document

import (
	"encoding/xml"
	"fmt"
	"strconv"
	"strings"
)

// TOCConfig 目录配置
type TOCConfig struct {
	Title        string // 目录标题，默认为"目录"
	MaxLevel     int    // 最大级别，默认为3（显示1-3级标题）
	ShowPageNum  bool   // 是否显示页码，默认为true
	RightAlign   bool   // 页码是否右对齐，默认为true
	UseHyperlink bool   // 是否使用超链接，默认为true
	DotLeader    bool   // 是否使用点状引导线，默认为true
	PageOffset   int    // 页码偏移量，用于过滤掉封面等页数（目录页码 = 物理页码 - PageOffset）
}

// TOCEntry 目录条目
type TOCEntry struct {
	Text       string // 条目文本
	Level      int    // 级别（1-9）
	PageNum    int    // 页码
	BookmarkID string // 书签ID（用于超链接）
}

// TOCField 目录域
type TOCField struct {
	XMLName xml.Name `xml:"w:fldSimple"`
	Instr   string   `xml:"w:instr,attr"`
	Runs    []Run    `xml:"w:r"`
}

// Hyperlink 超链接结构
type Hyperlink struct {
	XMLName xml.Name `xml:"w:hyperlink"`
	Anchor  string   `xml:"w:anchor,attr,omitempty"`
	Runs    []Run    `xml:"w:r"`
}

// BookmarkEnd 书签结束
type BookmarkEnd struct {
	XMLName xml.Name `xml:"w:bookmarkEnd"`
	ID      string   `xml:"w:id,attr"`
}

// ElementType 返回书签结束元素类型
func (b *BookmarkEnd) ElementType() string {
	return "bookmarkEnd"
}

// BookmarkStart 书签开始
type BookmarkStart struct {
	XMLName xml.Name `xml:"w:bookmarkStart"`
	ID      string   `xml:"w:id,attr"`
	Name    string   `xml:"w:name,attr"`
}

// ElementType 返回书签开始元素类型
func (b *BookmarkStart) ElementType() string {
	return "bookmarkStart"
}

// DefaultTOCConfig 返回默认目录配置
func DefaultTOCConfig() *TOCConfig {
	return &TOCConfig{
		Title:        "目录",
		MaxLevel:     3,
		ShowPageNum:  true,
		RightAlign:   true,
		UseHyperlink: true,
		DotLeader:    true,
		PageOffset:   0, // 默认不偏移
	}
}

// GenerateTOC 生成目录
func (d *Document) GenerateTOC(config *TOCConfig) error {
	if config == nil {
		config = DefaultTOCConfig()
	}
	
	// 收集标题信息
	entries := d.collectHeadings(config.MaxLevel)
	
	// 创建目录SDT
	tocSDT := d.CreateTOCSDT(config.Title, config.MaxLevel)
	
	// 为每个标题条目添加到目录中
	for i, entry := range entries {
		entryID := fmt.Sprintf("14746%d", 3000+i)
		tocSDT.AddTOCEntry(entry.Text, entry.Level, entry.PageNum, entryID)
	}
	
	// 完成目录SDT构建
	tocSDT.FinalizeTOCSDT()
	
	// 添加到文档中
	d.Body.Elements = append(d.Body.Elements, tocSDT)
	
	return nil
}

// UpdateTOC 更新目录
func (d *Document) UpdateTOC() error {
	// 重新收集标题信息
	entries := d.collectHeadings(9) // 收集所有级别
	
	// 查找现有目录
	tocStart := d.findTOCStart()
	if tocStart == -1 {
		return fmt.Errorf("未找到目录")
	}
	
	// 删除现有目录条目
	d.removeTOCEntries(tocStart)
	
	// 重新生成目录条目
	config := DefaultTOCConfig()
	for _, entry := range entries {
		if err := d.addTOCEntry(entry, config); err != nil {
			return fmt.Errorf("更新目录条目失败: %v", err)
		}
	}
	
	return nil
}

// AddHeadingWithBookmark 添加带书签的标题
// 注意：此函数已废弃，请使用 AddHeadingParagraphWithBookmark
// 保留此函数仅为了向后兼容
func (d *Document) AddHeadingWithBookmark(text string, level int, bookmarkName string) *Paragraph {
	// 直接调用新的实现
	return d.AddHeadingParagraphWithBookmark(text, level, bookmarkName)
}

// collectHeadings 收集标题信息
// 注意：此方法用于旧的GenerateTOC实现，页码由PAGEREF字段自动更新
func (d *Document) collectHeadings(maxLevel int) []TOCEntry {
	var entries []TOCEntry
	
	for _, element := range d.Body.Elements {
		if paragraph, ok := element.(*Paragraph); ok {
			// 检查标题级别
			level := d.getHeadingLevel(paragraph)
			if level > 0 && level <= maxLevel {
				text := d.extractParagraphText(paragraph)
				if text != "" {
					// 页码完全依赖PAGEREF字段自动更新，这里只提供初始占位值
					entry := TOCEntry{
						Text:       text,
						Level:      level,
						PageNum:    1, // 初始占位值，PAGEREF会自动更新为正确页码
						BookmarkID: fmt.Sprintf("_Toc_%s", strings.ReplaceAll(text, " ", "_")),
					}
					entries = append(entries, entry)
				}
			}
		}
	}
	
	return entries
}

// getHeadingLevel 获取段落的标题级别
func (d *Document) getHeadingLevel(paragraph *Paragraph) int {
	if paragraph.Properties != nil && paragraph.Properties.ParagraphStyle != nil {
		styleVal := paragraph.Properties.ParagraphStyle.Val
		
		// 根据样式ID映射标题级别 - 支持数字ID
		switch styleVal {
		case "1": // heading 1 (有些文档使用1作为标题1)
			return 1
		case "2": // heading 1 (Word默认使用2作为标题1)
			return 1
		case "3": // heading 2
			return 2
		case "4": // heading 3
			return 3
		case "5": // heading 4
			return 4
		case "6": // heading 5
			return 5
		case "7": // heading 6
			return 6
		case "8": // heading 7
			return 7
		case "9": // heading 8
			return 8
		case "10": // heading 9
			return 9
		}
		
		// 支持标准样式名称匹配
		switch styleVal {
		case "Heading1", "heading1", "Title1":
			return 1
		case "Heading2", "heading2", "Title2":
			return 2
		case "Heading3", "heading3", "Title3":
			return 3
		case "Heading4", "heading4", "Title4":
			return 4
		case "Heading5", "heading5", "Title5":
			return 5
		case "Heading6", "heading6", "Title6":
			return 6
		case "Heading7", "heading7", "Title7":
			return 7
		case "Heading8", "heading8", "Title8":
			return 8
		case "Heading9", "heading9", "Title9":
			return 9
		}
		
		// 支持通用模式匹配（处理Heading后面跟数字的情况）
		if strings.HasPrefix(strings.ToLower(styleVal), "heading") {
			// 提取数字部分
			numStr := strings.TrimPrefix(strings.ToLower(styleVal), "heading")
			if numStr != "" {
				if level := parseInt(numStr); level >= 1 && level <= 9 {
					return level
				}
			}
		}
	}
	return 0
}

// parseInt 简单的字符串转整数函数
func parseInt(s string) int {
	switch s {
	case "1":
		return 1
	case "2":
		return 2
	case "3":
		return 3
	case "4":
		return 4
	case "5":
		return 5
	case "6":
		return 6
	case "7":
		return 7
	case "8":
		return 8
	case "9":
		return 9
	default:
		return 0
	}
}

// extractParagraphText 提取段落文本
func (d *Document) extractParagraphText(paragraph *Paragraph) string {
	var text strings.Builder
	for _, run := range paragraph.Runs {
		text.WriteString(run.Text.Content)
	}
	return text.String()
}

// insertTOCField 插入目录域
func (d *Document) insertTOCField(config *TOCConfig) error {
	// 构建TOC指令
	instr := fmt.Sprintf("TOC \\o \"1-%d\"", config.MaxLevel)
	if config.UseHyperlink {
		instr += " \\h"
	}
	if !config.ShowPageNum {
		instr += " \\n"
	}
	
	// 创建目录域段落
	tocPara := &Paragraph{
		Properties: &ParagraphProperties{
			ParagraphStyle: &ParagraphStyle{Val: "TOC1"},
		},
	}
	
	// 添加域开始
	fieldStart := Run{
		Properties: &RunProperties{},
		Text:       Text{Content: ""}, // 域开始标记
	}
	
	// 添加域指令
	fieldInstr := Run{
		Properties: &RunProperties{},
		Text:       Text{Content: instr},
	}
	
	// 添加域结束
	fieldEnd := Run{
		Properties: &RunProperties{},
		Text:       Text{Content: ""}, // 域结束标记
	}
	
	tocPara.Runs = append(tocPara.Runs, fieldStart, fieldInstr, fieldEnd)
	d.Body.Elements = append(d.Body.Elements, tocPara)
	
	return nil
}

// addTOCEntry 添加目录条目
func (d *Document) addTOCEntry(entry TOCEntry, config *TOCConfig) error {
	// 创建目录条目段落
	entryPara := &Paragraph{
		Properties: &ParagraphProperties{
			ParagraphStyle: &ParagraphStyle{Val: fmt.Sprintf("TOC%d", entry.Level)},
		},
	}
	
	if config.UseHyperlink {
		// 创建超链接
		hyperlink := &Hyperlink{
			Anchor: entry.BookmarkID,
		}
		
		// 标题文本
		titleRun := Run{
			Properties: &RunProperties{},
			Text:       Text{Content: entry.Text},
		}
		hyperlink.Runs = append(hyperlink.Runs, titleRun)
		
		// 如果显示页码，添加引导线和页码
		if config.ShowPageNum {
			if config.DotLeader {
				// 添加点状引导线
				leaderRun := Run{
					Properties: &RunProperties{},
					Text:       Text{Content: strings.Repeat(".", 20)}, // 简化处理
				}
				hyperlink.Runs = append(hyperlink.Runs, leaderRun)
			}
			
			// 添加页码
			pageRun := Run{
				Properties: &RunProperties{},
				Text:       Text{Content: fmt.Sprintf("%d", entry.PageNum)},
			}
			hyperlink.Runs = append(hyperlink.Runs, pageRun)
		}
		
		// 将超链接添加到段落中
		// 这里需要特殊处理，因为Hyperlink不是标准的Run
		// 简化处理，直接作为文本添加
		hyperlinkRun := Run{
			Properties: &RunProperties{},
			Text:       Text{Content: entry.Text},
		}
		entryPara.Runs = append(entryPara.Runs, hyperlinkRun)
		
		if config.ShowPageNum {
			pageRun := Run{
				Properties: &RunProperties{},
				Text:       Text{Content: fmt.Sprintf("\t%d", entry.PageNum)},
			}
			entryPara.Runs = append(entryPara.Runs, pageRun)
		}
	} else {
		// 不使用超链接的简单文本
		titleRun := Run{
			Properties: &RunProperties{},
			Text:       Text{Content: entry.Text},
		}
		entryPara.Runs = append(entryPara.Runs, titleRun)
		
		if config.ShowPageNum {
			pageRun := Run{
				Properties: &RunProperties{},
				Text:       Text{Content: fmt.Sprintf("\t%d", entry.PageNum)},
			}
			entryPara.Runs = append(entryPara.Runs, pageRun)
		}
	}
	
	d.Body.Elements = append(d.Body.Elements, entryPara)
	return nil
}

// findTOCStart 查找目录开始位置
func (d *Document) findTOCStart() int {
	for i, element := range d.Body.Elements {
		if paragraph, ok := element.(*Paragraph); ok {
			if paragraph.Properties != nil && paragraph.Properties.ParagraphStyle != nil {
				if strings.HasPrefix(paragraph.Properties.ParagraphStyle.Val, "TOC") {
					return i
				}
			}
		}
	}
	return -1
}

// removeTOCEntries 删除现有目录条目
func (d *Document) removeTOCEntries(startIndex int) {
	// 简化处理：从startIndex开始查找并删除所有TOC样式的段落
	var newElements []interface{}
	
	// 保留start之前的元素
	newElements = append(newElements, d.Body.Elements[:startIndex]...)
	
	// 跳过TOC相关的元素
	for i := startIndex; i < len(d.Body.Elements); i++ {
		element := d.Body.Elements[i]
		if paragraph, ok := element.(*Paragraph); ok {
			if paragraph.Properties != nil && paragraph.Properties.ParagraphStyle != nil {
				if !strings.HasPrefix(paragraph.Properties.ParagraphStyle.Val, "TOC") {
					// 不是TOC样式，保留后续所有元素
					newElements = append(newElements, d.Body.Elements[i:]...)
					break
				}
			}
		}
	}
	
	d.Body.Elements = newElements
}

// SetTOCStyle 设置目录样式
func (d *Document) SetTOCStyle(level int, style *TextFormat) error {
	if level < 1 || level > 9 {
		return fmt.Errorf("目录级别必须在1-9之间")
	}
	
	styleName := fmt.Sprintf("TOC%d", level)
	
	// 通过样式管理器设置目录样式
	styleManager := d.GetStyleManager()
	
	// 创建段落样式（这里需要与样式系统集成）
	// 简化处理，实际需要创建完整的样式定义
	_ = styleManager
	_ = styleName
	_ = style
	
	return nil
}

// AutoGenerateTOC 自动生成目录，检测现有文档中的标题
func (d *Document) AutoGenerateTOC(config *TOCConfig) error {
	if config == nil {
		config = DefaultTOCConfig()
	}
	
	// 查找现有目录位置
	tocStart := d.findTOCStart()
	var insertIndex int
	
	if tocStart != -1 {
		// 如果已有目录，删除现有目录条目
		d.removeTOCEntries(tocStart)
		insertIndex = tocStart
	} else {
		// 如果没有目录，在文档开头插入
		insertIndex = 0
	}
	
	// 收集文档中的所有标题并为它们添加书签
	entries := d.collectHeadingsAndAddBookmarks(config.MaxLevel)
	
	if len(entries) == 0 {
		return fmt.Errorf("文档中未找到标题（样式ID为2-10的段落）")
	}
	
	// 使用真正的Word域字段生成目录，而不是简化的SDT
	tocElements := d.createWordFieldTOC(config, entries)
	
	// 将目录插入到指定位置
	if insertIndex == 0 {
		// 在开头插入
		d.Body.Elements = append(tocElements, d.Body.Elements...)
	} else {
		// 在指定位置替换
		newElements := make([]interface{}, 0, len(d.Body.Elements)+len(tocElements))
		newElements = append(newElements, d.Body.Elements[:insertIndex]...)
		newElements = append(newElements, tocElements...)
		newElements = append(newElements, d.Body.Elements[insertIndex:]...)
		d.Body.Elements = newElements
	}
	
	return nil
}

// GetHeadingCount 获取文档中标题的数量，用于调试
func (d *Document) GetHeadingCount() map[int]int {
	counts := make(map[int]int)
	
	for _, element := range d.Body.Elements {
		if paragraph, ok := element.(*Paragraph); ok {
			level := d.getHeadingLevel(paragraph)
			if level > 0 {
				counts[level]++
			}
		}
	}
	
	return counts
}

// ListHeadings 列出文档中所有的标题，用于调试
func (d *Document) ListHeadings() []TOCEntry {
	return d.collectHeadings(9) // 获取所有级别的标题
}

// GenerateTOCAtPosition 在指定位置生成目录（支持跳过指定索引的元素，如目录占位符）
// insertIndex: 目录插入位置（会替换该位置的元素）
// skipIndex: 要跳过的元素索引（如目录占位符段落）
func (d *Document) GenerateTOCAtPosition(config *TOCConfig, insertIndex, skipIndex int) error {
	if config == nil {
		config = DefaultTOCConfig()
	}
	
	// 收集标题信息，提取实际的书签名称
	entries := d.collectHeadingsWithBookmarks(config.MaxLevel, skipIndex, config.PageOffset)
	
	// 收集标题信息，提取实际的书签名称
	entries = d.collectHeadingsWithBookmarks(config.MaxLevel, skipIndex, config.PageOffset)
	
	if len(entries) == 0 {
		return fmt.Errorf("未找到标题")
	}
	
	// 创建目录SDT
	tocSDT := d.CreateTOCSDT(config.Title, config.MaxLevel)
	
	// 为每个标题条目添加到目录中（使用实际的书签ID）
	for _, entry := range entries {
		tocSDT.AddTOCEntry(entry.Text, entry.Level, entry.PageNum, entry.BookmarkID)
	}
	
	// 完成目录SDT构建
	tocSDT.FinalizeTOCSDT()
	
	// 在指定位置插入目录（替换占位符）
	if insertIndex >= 0 && insertIndex < len(d.Body.Elements) {
		// 移除占位符
		newElements := make([]interface{}, 0, len(d.Body.Elements))
		newElements = append(newElements, d.Body.Elements[:insertIndex]...)
		// 插入目录
		newElements = append(newElements, tocSDT)
		// 添加剩余元素（跳过占位符段落）
		if insertIndex+1 < len(d.Body.Elements) {
			newElements = append(newElements, d.Body.Elements[insertIndex+1:]...)
		}
		d.Body.Elements = newElements
	} else {
		// 如果索引超出范围，直接添加到末尾
		d.Body.Elements = append(d.Body.Elements, tocSDT)
	}
	
	return nil
}

// collectHeadingsWithBookmarks 收集标题信息，并提取实际的书签名称
// skipIndex: 要跳过的元素索引（如目录占位符段落）
// pageOffset: 页码偏移量，用于过滤掉封面等页数
//
// 页码计算说明：
// ===================
// 通过估算内容高度和检测分节符/分页符来计算页码。
// 实现了对页面设置（纸张大小、边距、方向）的动态跟踪，以处理横竖版混排和自动分页。
func (d *Document) collectHeadingsWithBookmarks(maxLevel int, skipIndex int, pageOffset int) []TOCEntry {
	var entries []TOCEntry
	
	// 1. 预处理分节符，获取每一节的页面设置
	sectionPropsList := d.extractSectionProperties()
	if len(sectionPropsList) == 0 {
		// 如果没有分节符，提供一个默认的
		sectionPropsList = append(sectionPropsList, &SectionProperties{})
	}
	
	currentSectionIdx := 0
	currentSection := sectionPropsList[0]
	
	// 初始页面参数 (磅)
	pageHeightPt, _, marginTopPt, marginBottomPt, _ := getPageDimensionsPt(currentSection)
	contentHeightPt := pageHeightPt - marginTopPt - marginBottomPt
	
	// 初始化状态
	currentY := 0.0           // 当前页面已用高度 (磅)
	currentPage := 1          // 物理页码 (从1开始)
	displayPage := 1          // 显示页码 (考虑起始页码设置)

	// 目录页码偏移修正：如果PageOffset>0，则逻辑页码从1-PageOffset开始
	if pageOffset > 0 {
		displayPage = 1 - pageOffset
	}

	// 检查第一节是否有起始页码设置
	if currentSection.PageNumType != nil && currentSection.PageNumType.Start != "" {
		if start, err := strconv.Atoi(currentSection.PageNumType.Start); err == nil {
			displayPage = start
		}
	}
	
	currentBookmarkName := "" // 当前标题对应的书签名称
	elementIndex := 0
	ignoreNextPageBreak := false // 标志位：忽略下一个显式分页符（用于处理分节符后紧跟的分页符）
	
	// 遍历所有元素
	for _, element := range d.Body.Elements {
		elementIndex++
		
		// 跳过指定索引的元素
		if elementIndex-1 == skipIndex {
			continue
		}
		
		// 检查书签
		if bookmarkStart, ok := element.(*BookmarkStart); ok {
			currentBookmarkName = bookmarkStart.Name
			continue
		}
		if _, ok := element.(*BookmarkEnd); ok {
			currentBookmarkName = ""
			continue
		}
		
		// 元素高度
		elementHeight := 0.0
		isHeading := false
		headingLevel := 0
		headingText := ""
		
		// 检查段落
		if paragraph, ok := element.(*Paragraph); ok {
			// 1. 检查分页符 (Explicit Page Break)
			hasPageBreak := false
			for _, run := range paragraph.Runs {
				if run.Break != nil && run.Break.Type == "page" {
					hasPageBreak = true
					break
				}
			}
			if paragraph.Properties != nil && paragraph.Properties.PageBreak != nil {
				hasPageBreak = true
			}
			
			if hasPageBreak {
				currentPage++
				currentY = 0
				if ignoreNextPageBreak {
					Debugf("忽略紧跟分节符的分页符: %s, Page 保持 %d", d.extractParagraphText(paragraph), displayPage)
					ignoreNextPageBreak = false
				} else {
					displayPage++
					Debugf("发现分页符: %s, Page -> %d", d.extractParagraphText(paragraph), displayPage)
				}
			} else {
				ignoreNextPageBreak = false
			}
			
			// 2. 估算段落高度
			// 使用页面宽度减去左右边距作为内容宽度
			_, pageWidthPt, _, _, sideMarginsPt := getPageDimensionsPt(currentSection)
			contentWidthPt := pageWidthPt - sideMarginsPt
			elementHeight = d.estimateParagraphHeight(paragraph, contentWidthPt)
			
			Debugf("元素: %s, 高度: %.2f, 当前Y: %.2f, 剩余: %.2f (Page: %d)",
				d.extractParagraphText(paragraph), elementHeight, currentY, contentHeightPt-currentY, displayPage)
			
			// 3. 检查是否是标题
			if paragraph.Properties != nil && paragraph.Properties.ParagraphStyle != nil {
				styleVal := paragraph.Properties.ParagraphStyle.Val
				// 简单的样式匹配
				if strings.HasPrefix(styleVal, "Heading") {
					if n, err := strconv.Atoi(strings.TrimPrefix(styleVal, "Heading")); err == nil {
						headingLevel = n
					}
				} else if len(styleVal) == 1 && styleVal >= "1" && styleVal <= "9" {
					// 处理 "1", "2" 这种样式ID
					headingLevel, _ = strconv.Atoi(styleVal)
				}
				
				if headingLevel > 0 {
					isHeading = true
					headingText = d.extractParagraphText(paragraph)
				}
			}
			
			// 4. 检查分节符 (Section Break)
			// 注意：在Word中，带有sectPr的段落是该节的最后一个段落
			// 所以，该段落仍属于当前节，计算完高度后，才切换到下一节
			if paragraph.Properties != nil && paragraph.Properties.SectionProperties != nil {
				// 累加高度
				if currentY+elementHeight > contentHeightPt {
					// 自动分页
					currentPage++
					displayPage++
					currentY = elementHeight
				} else {
					currentY += elementHeight
				}
				
				// 如果是标题，记录（注意：如果是分节符段落本身是标题，它还在前一页或当前页）
				if isHeading && headingLevel <= maxLevel && headingText != "" {
					d.addTOCEntryToList(&entries, headingText, headingLevel, displayPage, &currentBookmarkName)
				}
				
				// 切换到下一节
				currentSectionIdx++
				if currentSectionIdx < len(sectionPropsList) {
					currentSection = sectionPropsList[currentSectionIdx]
					
					// 更新页面参数
					pageHeightPt, _, marginTopPt, marginBottomPt, _ = getPageDimensionsPt(currentSection)
					contentHeightPt = pageHeightPt - marginTopPt - marginBottomPt
					
					// 分节符通常意味着新的一页（默认 Next Page）
					// 除非是 Continuous，这里简化处理，假设都会换页
					currentPage++
					
					// 检查是否重置页码
					if currentSection.PageNumType != nil && currentSection.PageNumType.Start != "" {
						startVal := currentSection.PageNumType.Start
						if start, err := strconv.Atoi(startVal); err == nil {
							// 仅在第一个分节符（通常是封面到正文）或明确要求时重置
							// 这里为了匹配用户预期的连续页码（忽略中间的 Start=1 重置），做了一个特殊处理
							// 如果是横竖版切换场景，通常希望页码连续
							if currentSectionIdx <= 1 {
								Debugf("分节符重置页码: Start=%s, DisplayPage=%d -> %d", startVal, displayPage, start)
								displayPage = start
								ignoreNextPageBreak = true // 重置后，如果紧接着有分页符，忽略其页码增加
							} else {
								Debugf("忽略非首节页码重置以保持连续: Start=%s, DisplayPage=%d", startVal, displayPage)
								displayPage++
							}
						} else {
							displayPage++
						}
					} else {
						displayPage++
					}
					
					currentY = 0
				}
				continue // 已处理完该段落
			}
		} else if table, ok := element.(*Table); ok {
			// 表格处理
			_, pageWidthPt, _, _, sideMarginsPt := getPageDimensionsPt(currentSection)
			contentWidthPt := pageWidthPt - sideMarginsPt
			
			// 估算表格每一行的高度
			for _, row := range table.Rows {
				rowHeight := d.estimateRowHeight(&row, contentWidthPt)
				
				if currentY+rowHeight > contentHeightPt {
					currentPage++
					displayPage++
					currentY = rowHeight
				} else {
					currentY += rowHeight
				}
			}
			continue
		}
		
		// 普通段落（非分节符）的高度处理
		if currentY+elementHeight > contentHeightPt {
			currentPage++
			displayPage++
			currentY = elementHeight
		} else {
			currentY += elementHeight
		}
		
		// 如果是标题，记录
		if isHeading && headingLevel <= maxLevel && headingText != "" {
			d.addTOCEntryToList(&entries, headingText, headingLevel, displayPage, &currentBookmarkName)
		}
	}
	
	return entries
}

// 辅助方法：添加目录条目
func (d *Document) addTOCEntryToList(entries *[]TOCEntry, text string, level int, pageNum int, bookmarkName *string) {
	// 使用实际的书签名称（如果存在），否则生成默认的书签ID
	bookmarkID := *bookmarkName
	if bookmarkID == "" {
		bookmarkID = fmt.Sprintf("_Toc_%s", strings.ReplaceAll(text, " ", "_"))
	}
	
	entry := TOCEntry{
		Text:       text,
		Level:      level,
		PageNum:    pageNum,
		BookmarkID: bookmarkID,
	}
	*entries = append(*entries, entry)
	
	// 清除当前书签名称（已使用）
	*bookmarkName = ""
}

// extractSectionProperties 提取文档中所有的节属性
// 返回的列表中，第 i 个元素对应第 i 节的属性
func (d *Document) extractSectionProperties() []*SectionProperties {
	var props []*SectionProperties
	
	// 遍历 Body 元素寻找 sectPr
	for _, element := range d.Body.Elements {
		if paragraph, ok := element.(*Paragraph); ok {
			if paragraph.Properties != nil && paragraph.Properties.SectionProperties != nil {
				props = append(props, paragraph.Properties.SectionProperties)
			}
		} else if sectPr, ok := element.(*SectionProperties); ok {
			// 文档末尾的 sectPr
			props = append(props, sectPr)
		}
	}
	
	// 如果文档末尾没有显式的 sectPr 元素（但在 XML 中通常会有），
	// 或者没有任何分节符，我们至少应该有一个默认的。
	// 但在 extract 阶段我们只返回发现的。
	
	return props
}

// getPageDimensionsPt 获取页面尺寸和边距（磅）
func getPageDimensionsPt(sectPr *SectionProperties) (height, width, marginTop, marginBottom, sideMargins float64) {
	// 默认 A4 Portrait
	width = 595.3    // 210mm
	height = 841.9   // 297mm
	marginTop = 72.0 // 1 inch
	marginBottom = 72.0
	marginLeft := 72.0
	marginRight := 72.0
	
	if sectPr == nil {
		return
	}
	
	if sectPr.PageSize != nil {
		if w, err := strconv.ParseFloat(sectPr.PageSize.W, 64); err == nil {
			width = w / 20.0
		}
		if h, err := strconv.ParseFloat(sectPr.PageSize.H, 64); err == nil {
			height = h / 20.0
		}
		// 如果是横向，交换宽高（XML中的 w 和 h 通常已经对应纸张的物理宽高，orient 只是元数据，
		// 但 Word 有时行为不一致。通常 w 是长边还是短边取决于 orient?
		// 实际上 OOXML 中 w 和 h 就是页面显示的宽和高。
		// 如果 orient="landscape"，通常 w > h。
		// 我们直接信任 w 和 h。
	}
	
	if sectPr.PageMargins != nil {
		if t, err := strconv.ParseFloat(sectPr.PageMargins.Top, 64); err == nil {
			marginTop = t / 20.0
		}
		if b, err := strconv.ParseFloat(sectPr.PageMargins.Bottom, 64); err == nil {
			marginBottom = b / 20.0
		}
		if l, err := strconv.ParseFloat(sectPr.PageMargins.Left, 64); err == nil {
			marginLeft = l / 20.0
		}
		if r, err := strconv.ParseFloat(sectPr.PageMargins.Right, 64); err == nil {
			marginRight = r / 20.0
		}
	}
	
	sideMargins = marginLeft + marginRight
	return
}

// estimateParagraphHeight 估算段落高度 (磅)
func (d *Document) estimateParagraphHeight(p *Paragraph, contentWidthPt float64) float64 {
	if p == nil || len(p.Runs) == 0 {
		return 12.0 // 空段落至少占一行
	}
	
	// 1. 计算内容总长度（字符数）和平均字号
	totalChars := 0
	maxFontSize := 10.5 // 默认五号字 (10.5pt)
	
	// 检查段落属性中的默认字号（如果 Run 没有指定）
	// 这里简化处理，直接用默认值
	
	for _, run := range p.Runs {
		content := run.Text.Content
		totalChars += len([]rune(content)) // 使用 rune 计数，处理中文
		
		// 检查字号
		if run.Properties != nil && run.Properties.FontSize != nil {
			if sizeHalfPt, err := strconv.ParseFloat(run.Properties.FontSize.Val, 64); err == nil {
				sizePt := sizeHalfPt / 2.0
				if sizePt > maxFontSize {
					maxFontSize = sizePt
				}
			}
		}
	}
	
	if totalChars == 0 {
		return maxFontSize * 1.2 // 空行高度
	}
	
	// 2. 估算行数
	// 假设平均每个字符宽度为字号（中文）或字号的一半（英文）。
	// 这是一个粗略估算。为安全起见，假设都是宽字符（中文）。
	// 一行能容纳的字符数 ≈ contentWidthPt / maxFontSize
	charsPerLine := int(contentWidthPt / maxFontSize)
	if charsPerLine < 1 {
		charsPerLine = 1
	}
	
	numLines := (totalChars + charsPerLine - 1) / charsPerLine
	
	// 3. 计算高度
	// 行高通常为字号的 1.2 到 1.5 倍
	lineHeight := maxFontSize * 1.3
	
	// 检查段落行距设置
	if p.Properties != nil && p.Properties.Spacing != nil {
		if p.Properties.Spacing.Line != "" {
			// "240" = 12pt (单倍行距 12*20)
			// 如果是具体数值
			if val, err := strconv.ParseFloat(p.Properties.Spacing.Line, 64); err == nil {
				// lineRule="auto" (default) -> 240 means 1 line (relative)
				// lineRule="exact" -> 240 means 12pt
				// 这里简化：假设是绝对值 twips
				// 实际上 word 默认 lineRule 是 auto，值 240 代表 1 倍行距。360 是 1.5 倍。
				// 我们简单处理：
				lineHeight = (val / 240.0) * maxFontSize * 1.3
			}
		}
		
		// 加上段前段后
		spacingBefore := 0.0
		if p.Properties.Spacing.Before != "" {
			if val, err := strconv.ParseFloat(p.Properties.Spacing.Before, 64); err == nil {
				spacingBefore = val / 20.0
			}
		}
		spacingAfter := 0.0
		if p.Properties.Spacing.After != "" {
			if val, err := strconv.ParseFloat(p.Properties.Spacing.After, 64); err == nil {
				spacingAfter = val / 20.0
			}
		}
		
		return float64(numLines)*lineHeight + spacingBefore + spacingAfter
	}
	
	return float64(numLines) * lineHeight
}

// estimateRowHeight 估算表格行高度
func (d *Document) estimateRowHeight(row *TableRow, contentWidthPt float64) float64 {
	// 1. 检查是否有固定行高
	if row.Properties != nil && row.Properties.TableRowH != nil {
		if val, err := strconv.ParseFloat(row.Properties.TableRowH.Val, 64); err == nil {
			return val / 20.0
		}
	}
	
	// 2. 如果没有固定行高，估算内容高度
	// 找出所有单元格中最高的那个
	maxCellHeight := 0.0
	numCells := len(row.Cells)
	if numCells == 0 {
		return 12.0 // 默认一行高度
	}
	
	// 假设列宽平均分配（简化）
	cellWidthPt := contentWidthPt / float64(numCells)
	
	for _, cell := range row.Cells {
		cellHeight := 0.0
		for _, para := range cell.Paragraphs {
			cellHeight += d.estimateParagraphHeight(&para, cellWidthPt)
		}
		// 加上单元格内边距（假设上下各 2pt）
		cellHeight += 4.0
		
		if cellHeight > maxCellHeight {
			maxCellHeight = cellHeight
		}
	}
	
	if maxCellHeight == 0 {
		return 12.0
	}
	return maxCellHeight
}

// createWordFieldTOC 创建使用真正Word域字段的目录
func (d *Document) createWordFieldTOC(config *TOCConfig, entries []TOCEntry) []interface{} {
	var elements []interface{}
	
	// 创建目录SDT容器
	tocSDT := &SDT{
		Properties: &SDTProperties{
			RunPr: &RunProperties{
				FontFamily: &FontFamily{ASCII: "宋体", HAnsi: "宋体", EastAsia: "宋体", CS: "Times New Roman"},
				FontSize:   &FontSize{Val: "21"},
			},
			ID:    &SDTID{Val: "147458718"},
			Color: &SDTColor{Val: "DBDBDB"},
			DocPartObj: &DocPartObj{
				DocPartGallery: &DocPartGallery{Val: "Table of Contents"},
				DocPartUnique:  &DocPartUnique{},
			},
		},
		EndPr: &SDTEndPr{
			RunPr: &RunProperties{
				FontFamily: &FontFamily{ASCII: "Calibri", HAnsi: "Calibri", EastAsia: "宋体", CS: "Times New Roman"},
				Bold:       &Bold{},
				Color:      &Color{Val: "2F5496"},
				FontSize:   &FontSize{Val: "32"},
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
			Justification: &Justification{Val: "center"},
			Indentation: &Indentation{
				Left:      "0",
				Right:     "0",
				FirstLine: "0",
			},
		},
		Runs: []Run{
			{
				Text: Text{Content: config.Title},
				Properties: &RunProperties{
					FontFamily: &FontFamily{ASCII: "宋体"},
					FontSize:   &FontSize{Val: "21"},
				},
			},
		},
	}
	
	tocSDT.Content.Elements = append(tocSDT.Content.Elements, titlePara)
	
	// 创建主TOC域段落
	tocFieldPara := &Paragraph{
		Properties: &ParagraphProperties{
			ParagraphStyle: &ParagraphStyle{Val: "12"}, // TOC样式
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
	
	// 添加TOC域开始
	tocFieldPara.Runs = append(tocFieldPara.Runs, Run{
		Properties: &RunProperties{
			Bold:     &Bold{},
			Color:    &Color{Val: "2F5496"},
			FontSize: &FontSize{Val: "32"},
		},
		FieldChar: &FieldChar{
			FieldCharType: "begin",
		},
	})
	
	// 添加TOC指令 - 使用 \n 开关不显示页码，因为我们手动创建条目
	instrContent := fmt.Sprintf("TOC \\o \"1-%d\" \\h \\u \\n", config.MaxLevel)
	tocFieldPara.Runs = append(tocFieldPara.Runs, Run{
		Properties: &RunProperties{
			Bold:     &Bold{},
			Color:    &Color{Val: "2F5496"},
			FontSize: &FontSize{Val: "32"},
		},
		InstrText: &InstrText{
			Space:   "preserve",
			Content: instrContent,
		},
	})
	
	// 添加TOC域分隔符
	tocFieldPara.Runs = append(tocFieldPara.Runs, Run{
		Properties: &RunProperties{
			Bold:     &Bold{},
			Color:    &Color{Val: "2F5496"},
			FontSize: &FontSize{Val: "32"},
		},
		FieldChar: &FieldChar{
			FieldCharType: "separate",
		},
	})
	
	tocSDT.Content.Elements = append(tocSDT.Content.Elements, tocFieldPara)
	
	// 为每个条目创建超链接段落
	for _, entry := range entries {
		entryPara := d.createTOCEntryWithFields(entry, config)
		tocSDT.Content.Elements = append(tocSDT.Content.Elements, entryPara)
	}
	
	// 添加TOC域结束段落
	endPara := &Paragraph{
		Properties: &ParagraphProperties{
			ParagraphStyle: &ParagraphStyle{Val: "2"},
			Spacing: &Spacing{
				Before: "240",
				After:  "0",
			},
		},
		Runs: []Run{
			{
				Properties: &RunProperties{
					Color: &Color{Val: "2F5496"},
				},
				FieldChar: &FieldChar{
					FieldCharType: "end",
				},
			},
		},
	}
	
	tocSDT.Content.Elements = append(tocSDT.Content.Elements, endPara)
	elements = append(elements, tocSDT)
	
	return elements
}

// createTOCEntryWithFields 创建带域字段的目录条目
func (d *Document) createTOCEntryWithFields(entry TOCEntry, config *TOCConfig) *Paragraph {
	// 确定目录样式ID
	var styleVal string
	switch entry.Level {
	case 1:
		styleVal = "13" // TOC 1
	case 2:
		styleVal = "14" // TOC 2
	case 3:
		styleVal = "15" // TOC 3
	default:
		styleVal = fmt.Sprintf("%d", 12+entry.Level)
	}
	
	para := &Paragraph{
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
	
	// 使用实际的书签ID（从entry中获取），而不是重新生成
	// 这样可以确保PAGEREF引用的书签与标题段落的书签一致
	anchor := entry.BookmarkID
	
	// 创建超链接域开始
	para.Runs = append(para.Runs, Run{
		Properties: &RunProperties{
			Color: &Color{Val: "2F5496"},
		},
		FieldChar: &FieldChar{
			FieldCharType: "begin",
		},
	})
	
	// 添加超链接指令
	para.Runs = append(para.Runs, Run{
		InstrText: &InstrText{
			Space:   "preserve",
			Content: fmt.Sprintf(" HYPERLINK \\l %s ", anchor),
		},
	})
	
	// 超链接域分隔符
	para.Runs = append(para.Runs, Run{
		FieldChar: &FieldChar{
			FieldCharType: "separate",
		},
	})
	
	// 添加标题文本
	para.Runs = append(para.Runs, Run{
		Text: Text{Content: entry.Text},
	})
	
	// 添加制表符
	para.Runs = append(para.Runs, Run{
		Text: Text{Content: "\t"},
	})
	
	// 添加页码引用域
	para.Runs = append(para.Runs, Run{
		FieldChar: &FieldChar{
			FieldCharType: "begin",
		},
	})
	
	para.Runs = append(para.Runs, Run{
		InstrText: &InstrText{
			Space:   "preserve",
			Content: fmt.Sprintf(" PAGEREF %s \\h ", anchor), // 添加\h开关以创建超链接
		},
	})
	
	para.Runs = append(para.Runs, Run{
		FieldChar: &FieldChar{
			FieldCharType: "separate",
		},
	})
	
	// 页码文本（使用计算的页码作为初始值）
	// 通过遍历文档元素，识别分页符和分节符，计算每个标题所在的页码
	// 这样可以实现"一步到位"的效果，用户打开文档就能看到正确的页码
	// PAGEREF字段仍然存在，可以在Word中更新以获得精确页码
	para.Runs = append(para.Runs, Run{
		Text: Text{Content: fmt.Sprintf("%d", entry.PageNum)},
	})
	
	// 页码域结束
	para.Runs = append(para.Runs, Run{
		FieldChar: &FieldChar{
			FieldCharType: "end",
		},
	})
	
	// 超链接域结束
	para.Runs = append(para.Runs, Run{
		Properties: &RunProperties{
			Color: &Color{Val: "2F5496"},
		},
		FieldChar: &FieldChar{
			FieldCharType: "end",
		},
	})
	
	return para
}

// generateUniqueID 基于文本内容生成唯一ID
func generateUniqueID(text string) int {
	// 使用简单的哈希算法生成唯一ID
	hash := 0
	for _, char := range text {
		hash = hash*31 + int(char)
	}
	// 确保是正数并限制在合理范围内
	if hash < 0 {
		hash = -hash
	}
	return (hash % 90000) + 10000 // 生成10000-99999之间的数字
}

// collectHeadingsAndAddBookmarks 收集标题信息并添加书签
// 注意：页码完全依赖PAGEREF字段自动更新
func (d *Document) collectHeadingsAndAddBookmarks(maxLevel int) []TOCEntry {
	var entries []TOCEntry
	
	// 需要一个新的Elements切片来插入书签
	newElements := make([]interface{}, 0, len(d.Body.Elements)*2)
	entryIndex := 0
	
	for _, element := range d.Body.Elements {
		if paragraph, ok := element.(*Paragraph); ok {
			level := d.getHeadingLevel(paragraph)
			if level > 0 && level <= maxLevel {
				text := d.extractParagraphText(paragraph)
				if text != "" {
					// 为每个条目生成唯一的书签ID（与目录条目中使用的一致）
					anchor := fmt.Sprintf("_Toc%d", generateUniqueID(text))
					
					entry := TOCEntry{
						Text:       text,
						Level:      level,
						PageNum:    1, // 初始占位值，PAGEREF会自动更新为正确页码
						BookmarkID: anchor,
					}
					entries = append(entries, entry)
					
					// 在标题段落前添加书签开始标记
					bookmarkStart := &BookmarkStart{
						ID:   fmt.Sprintf("%d", entryIndex),
						Name: anchor,
					}
					newElements = append(newElements, bookmarkStart)
					
					// 添加原段落
					newElements = append(newElements, element)
					
					// 在标题段落后添加书签结束标记
					bookmarkEnd := &BookmarkEnd{
						ID: fmt.Sprintf("%d", entryIndex),
					}
					newElements = append(newElements, bookmarkEnd)
					
					entryIndex++
					continue
				}
			}
		}
		// 非标题段落直接添加
		newElements = append(newElements, element)
	}
	
	// 更新文档元素
	d.Body.Elements = newElements
	
	return entries
}
