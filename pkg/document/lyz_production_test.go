package document

import (
	"testing"
	
	"github.com/ZeroHawkeye/wordZero/pkg/style"
)

// Mock functions to simulate content generation

func reportCoverMock(doc *Document) {
	doc.AddParagraph("封面")
	doc.AddParagraph("这是一个模拟的封面。")
	doc.AddParagraph("").AddPageBreak()
}

func reportExplainMock(doc *Document, textFormat *TextFormat, spacingConfig *SpacingConfig) {
	doc.AddParagraph("检测报告说明")
	doc.AddParagraph("这是一个模拟的检测报告说明。")
	doc.AddParagraph("").AddPageBreak()
}

func signPageMock(doc *Document, textFormat *TextFormat, spacingConfig *SpacingConfig) {
	doc.AddParagraph("签名页")
	doc.AddParagraph("这是一个模拟的签名页。")
	doc.AddParagraph("").AddPageBreak()
}

func shuYuAndDingYi1Mock(doc *Document, textFormat *TextFormat, spacingConfig *SpacingConfig) {
	bookmarkName := "_Toc_术语和定义"
	doc.AddHeadingParagraphWithBookmark("术语和定义", 1, bookmarkName, textFormat)
	doc.AddParagraph("这里是术语和定义的详细内容。")
	doc.AddParagraph("").AddPageBreak()
}

// TestProductionIssue reproduces the user's issue with page numbers and TOC
func TestProductionIssue(t *testing.T) {
	// 启用调试日志
	SetGlobalLevel(LogLevelDebug)
	
	doc := New()
	doc.SetPageSettings(&PageSettings{
		MarginTop:      25,
		MarginRight:    20,
		MarginBottom:   25,
		MarginLeft:     20,
		Orientation:    OrientationPortrait,
		Size:           PageSizeA4,
		HeaderDistance: 10, // 增加页眉距离，避免顶着顶部
		FooterDistance: 10, // 增加页脚距离，避免压着底部
	})
	
	textFormat := &TextFormat{
		FontFamily: "SimSun",
		FontSize:   12,
		FontColor:  "000000",
		Bold:       false,
	}
	spacingConfig := &SpacingConfig{}
	
	// 1. 封面
	reportCoverMock(doc)
	
	// 2. 检测报告说明
	reportExplainMock(doc, textFormat, spacingConfig)
	
	// 3. 签名页
	signPageMock(doc, textFormat, spacingConfig)
	// 4. 目录
	tocTitlePara := doc.AddParagraph("目录")
	tocTitlePara.SetStyle(style.StyleHeading1)
	tocInsertIndex := len(doc.Body.Elements) - 1
	
	// 目录通常单独一页
	//doc.AddParagraph("").AddPageBreak()
	
	// 添加分节符，准备开始正文
	// 从这里开始，我们需要页码从1开始
	sectionBreak1 := doc.AddParagraph("")
	sectionBreak1.AddSectionBreakWithStartPage(OrientationPortrait, doc, 1, false) // false 表示不继承上一节（无）的页眉页脚
	// 在新节开始处设置页眉页脚
	err := doc.AddStyleHeader(HeaderFooterTypeDefault, "xxx科技有限公司 RLHB", "", &TextFormat{
		FontFamily: "SimSun",
		FontSize:   9,
		FontColor:  "000000",
	})
	if err != nil {
		t.Error(err)
	}
	
	if err := doc.AddFooterWithPageNumber(HeaderFooterTypeDefault, "", true); err != nil {
		t.Error(err)
	}
	
	// 5. 术语和定义 (正文第1页)
	shuYuAndDingYi1Mock(doc, textFormat, spacingConfig)
	
	// 6. 横版内容测试
	sectionBreak2 := doc.AddParagraph("")
	sectionBreak2.AddSectionBreakWithStartPage(OrientationLandscape, doc, 0, true) // 0=延续页码, true=继承页眉页脚
	
	bookmarkName2 := "_Toc_横版测试"
	doc.AddHeadingParagraphWithBookmark("横版测试标题", 1, bookmarkName2, textFormat)
	doc.AddParagraph("这是横版页面的内容，页码应该连续。")
	//doc.AddParagraph("").AddPageBreak()
	
	// 7. 切回竖版
	sectionBreak3 := doc.AddParagraph("")
	sectionBreak3.AddSectionBreakWithStartPage(OrientationPortrait, doc, 0, true)
	
	bookmarkName3 := "_Toc_竖版回归"
	doc.AddHeadingParagraphWithBookmark("竖版回归标题", 1, bookmarkName3, textFormat)
	doc.AddParagraph("这是回归竖版后的内容，页码应该继续连续。")
	
	// 生成目录
	tocConfig := &TOCConfig{
		Title:        "目录",
		MaxLevel:     3,
		ShowPageNum:  true,
		RightAlign:   true,
		UseHyperlink: true,
		DotLeader:    true,
		PageOffset:   0, // 目录页码偏移量设为0，确保逻辑页码正确
	}
	
	if err := doc.GenerateTOCAtPosition(tocConfig, tocInsertIndex, tocInsertIndex); err != nil {
		t.Error(err)
	}
	
	outputPath := "lyz_production_test.docx"
	if err := doc.Save(outputPath); err != nil {
		t.Error(err)
		return
	}
	t.Logf("文档已保存: %s", outputPath)
}
