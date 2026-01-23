// @Author: liyongzhen
// @Description: 修复后的页码测试 - 竖版横版交叉，页码连续
// @File: lyz_test_fixed
// @Date: 2025/6/19 16:10

package document

import (
	"testing"
	
	"github.com/ZeroHawkeye/wordZero/pkg/style"
)

func TestHeaderStyleFixed(t *testing.T) {
	// 启用调试日志
	SetGlobalLevel(LogLevelDebug)
	
	doc := New()
	doc.SetPageSettings(&PageSettings{
		MarginTop:    25,
		MarginRight:  20,
		MarginBottom: 25,
		MarginLeft:   20,
		Orientation:  OrientationPortrait,
		Size:         PageSizeA4,
	})
	
	// 首页 - 只创建内容，不设置页眉页脚
	coverTitle := doc.AddParagraph("文档标题")
	coverTitle.SetStyle(style.StyleTitle)
	coverTitle.SetAlignment(AlignCenter)
	
	coverContent := doc.AddParagraph("公司名称")
	coverContent.SetAlignment(AlignCenter)
	textFormat := &TextFormat{
		FontFamily: "SimSun",
		FontSize:   14,
		FontColor:  "000000",
		Bold:       true,
	}
	reportExplain(doc, textFormat, &SpacingConfig{})
	//reportExplain(doc, textFormat, &SpacingConfig{})
	
	doc.AddParagraph("").AddPageBreak()
	
	// 目录页 - 先创建目录标题占位符，后续会被目录内容替换
	tocTitlePara := doc.AddParagraph("目录")
	tocTitlePara.SetStyle(style.StyleHeading1)
	
	// 记录目录插入位置（占位符段落的位置）
	tocInsertIndex := len(doc.Body.Elements) - 1
	// 添加分节符，从这里开始启用新的节（竖版）
	// 使用新的 AddSectionBreakWithStartPage 方法，设置起始页码为1，并继承页眉页脚
	// 注意：这里不设置页眉页脚，继续添加内容直到第5页
	sectionBreak1 := doc.AddParagraph("")
	sectionBreak1.AddSectionBreakWithStartPage(OrientationPortrait, doc, 1, false)
	// 目录之后，应该是第1页
	shuban(doc, textFormat)
	
	// 为第1-4页添加带书签的标题段落（不显示页眉页脚）
	bookmarkName1 := "_Toc_第1页内容"
	pb1 := doc.AddHeadingParagraphWithBookmark("第1页内容", 1, bookmarkName1)
	// 添加第1页的详细内容
	content1 := doc.AddParagraph("这是第1页的详细内容。这里可以放置一些介绍性文字，说明文档的背景、目的和主要内容。")
	content1.SetStyle(style.StyleNormal)
	content1 = doc.AddParagraph("第1页还可以包含更多的段落，用于详细阐述相关内容。")
	content1.SetStyle(style.StyleNormal)
	pb1.AddPageBreak()
	
	bookmarkName2 := "_Toc_第2页内容"
	pb2 := doc.AddHeadingParagraphWithBookmark("第2页内容", 1, bookmarkName2)
	// 添加第2页的详细内容
	content2 := doc.AddParagraph("这是第2页的详细内容。可以继续扩展文档的内容，提供更多的信息和细节。")
	content2.SetStyle(style.StyleNormal)
	content2 = doc.AddParagraph("第2页的内容可以与第1页形成逻辑上的连贯性。")
	content2.SetStyle(style.StyleNormal)
	pb2.AddPageBreak()
	
	bookmarkName3 := "_Toc_第3页内容"
	pb3 := doc.AddHeadingParagraphWithBookmark("第3页内容", 1, bookmarkName3)
	// 添加第3页的详细内容
	content3 := doc.AddParagraph("这是第3页的详细内容。文档的中间部分通常包含核心内容和分析。")
	content3.SetStyle(style.StyleNormal)
	content3 = doc.AddParagraph("第3页可以包含数据、图表说明或其他重要信息。")
	content3.SetStyle(style.StyleNormal)
	pb3.AddPageBreak()
	
	bookmarkName4 := "_Toc_第4页内容"
	pb4 := doc.AddHeadingParagraphWithBookmark("第4页内容", 1, bookmarkName4)
	// 添加第4页的详细内容
	content4 := doc.AddParagraph("这是第4页的详细内容。接近文档的结尾部分，可以开始总结和归纳。")
	content4.SetStyle(style.StyleNormal)
	content4 = doc.AddParagraph("第4页为后续的正式内容做好铺垫。")
	content4.SetStyle(style.StyleNormal)
	pb4.AddPageBreak()
	
	// 添加分节符，从第5页开始显示页码，页码从1开始
	sectionBreak2 := doc.AddParagraph("")
	sectionBreak2.AddSectionBreakWithStartPage(OrientationPortrait, doc, 1, true)
	
	// 在新节上配置页眉页脚（只影响新节及之后内容）
	err := doc.AddStyleHeader(HeaderFooterTypeDefault, "xxx科技有限公司\nRLHB", "2025010", &TextFormat{
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
	// 使用 AddHeadingParagraphWithBookmark 创建标题，这样可以被目录生成功能识别并支持跳转
	// 一级标题：四号（14磅），宋体，加粗
	bookmarkName5 := "_Toc_第五页标题"
	p := doc.AddHeadingParagraphWithBookmark("第五页标题", 1, bookmarkName5)
	if len(p.Runs) > 0 {
		if p.Runs[0].Properties == nil {
			p.Runs[0].Properties = &RunProperties{}
		}
		// 设置字体格式：四号（14磅 * 2 = 28）
		p.Runs[0].Properties.FontFamily = &FontFamily{ASCII: "SimSun", HAnsi: "SimSun", EastAsia: "SimSun"}
		p.Runs[0].Properties.FontSize = &FontSize{Val: "28"}
		p.Runs[0].Properties.Color = &Color{Val: "000000"}
		p.Runs[0].Properties.Bold = &Bold{}
	}
	
	// 添加二级标题：小四（12磅），宋体，加粗
	bookmarkName5_1 := "_Toc_第五页二级标题1"
	p5_1 := doc.AddHeadingParagraphWithBookmark("第五页二级标题1", 2, bookmarkName5_1)
	if len(p5_1.Runs) > 0 {
		if p5_1.Runs[0].Properties == nil {
			p5_1.Runs[0].Properties = &RunProperties{}
		}
		// 设置字体格式：小四（12磅 * 2 = 24）
		p5_1.Runs[0].Properties.FontFamily = &FontFamily{ASCII: "SimSun", HAnsi: "SimSun", EastAsia: "SimSun"}
		p5_1.Runs[0].Properties.FontSize = &FontSize{Val: "24"}
		p5_1.Runs[0].Properties.Color = &Color{Val: "000000"}
		p5_1.Runs[0].Properties.Bold = &Bold{}
	}
	
	// 添加三级标题：小四（12磅），宋体，加粗
	bookmarkName5_1_1 := "_Toc_第五页三级标题1"
	p5_1_1 := doc.AddHeadingParagraphWithBookmark("第五页三级标题1", 3, bookmarkName5_1_1)
	if len(p5_1_1.Runs) > 0 {
		if p5_1_1.Runs[0].Properties == nil {
			p5_1_1.Runs[0].Properties = &RunProperties{}
		}
		// 设置字体格式：小四（12磅 * 2 = 24）
		p5_1_1.Runs[0].Properties.FontFamily = &FontFamily{ASCII: "SimSun", HAnsi: "SimSun", EastAsia: "SimSun"}
		p5_1_1.Runs[0].Properties.FontSize = &FontSize{Val: "24"}
		p5_1_1.Runs[0].Properties.Color = &Color{Val: "000000"}
		p5_1_1.Runs[0].Properties.Bold = &Bold{}
	}
	
	// 添加另一个二级标题：小四（12磅），宋体，加粗
	bookmarkName5_2 := "_Toc_第五页二级标题2"
	p5_2 := doc.AddHeadingParagraphWithBookmark("第五页二级标题2", 2, bookmarkName5_2)
	if len(p5_2.Runs) > 0 {
		if p5_2.Runs[0].Properties == nil {
			p5_2.Runs[0].Properties = &RunProperties{}
		}
		// 设置字体格式：小四（12磅 * 2 = 24）
		p5_2.Runs[0].Properties.FontFamily = &FontFamily{ASCII: "SimSun", HAnsi: "SimSun", EastAsia: "SimSun"}
		p5_2.Runs[0].Properties.FontSize = &FontSize{Val: "24"}
		p5_2.Runs[0].Properties.Color = &Color{Val: "000000"}
		p5_2.Runs[0].Properties.Bold = &Bold{}
	}
	
	// 在标题后添加一些内容，确保页面有内容
	contentPara := doc.AddParagraph("这是第五页标题页面的内容，包含二级标题和三级标题的测试。")
	if len(contentPara.Runs) > 0 {
		if contentPara.Runs[0].Properties == nil {
			contentPara.Runs[0].Properties = &RunProperties{}
		}
		contentPara.Runs[0].Properties.FontFamily = &FontFamily{ASCII: "SimSun", HAnsi: "SimSun", EastAsia: "SimSun"}
		contentPara.Runs[0].Properties.FontSize = &FontSize{Val: "24"} // 12磅 * 2
		contentPara.Runs[0].Properties.Color = &Color{Val: "000000"}
	}
	
	// 在内容段落后添加分页符和分节符（切换到横版）
	contentPara.AddPageBreak()
	contentPara.SetSpacing(&SpacingConfig{
		BeforePara: 0,
		AfterPara:  0,
	})
	
	// 使用 AddSectionBreakWithStartPage 保持页码连续，并继承页眉页脚
	// startPage=0 表示延续上一节的页码
	// inheritHeaderFooter=true 表示继承上一节的页眉页脚
	contentPara.AddSectionBreakWithStartPage(OrientationLandscape, doc, 0, true)
	
	// 如果需要自定义格式，可以更新Run的属性
	if len(p.Runs) > 0 {
		if p.Runs[0].Properties == nil {
			p.Runs[0].Properties = &RunProperties{}
		}
		// 设置字体格式
		p.Runs[0].Properties.FontFamily = &FontFamily{ASCII: "SimSun", HAnsi: "SimSun", EastAsia: "SimSun"}
		p.Runs[0].Properties.FontSize = &FontSize{Val: "28"} // 14磅 * 2
		p.Runs[0].Properties.Color = &Color{Val: "000000"}
		p.Runs[0].Properties.Bold = &Bold{}
	}
	
	textFormat.Bold = false
	textFormat.FontSize = 12
	// 表格:标准依据
	tableBz := doc.AddTable(&TableConfig{
		Rows:      23,
		Cols:      3,
		ColWidths: []int{2000, 4000, 3000},
		Width:     9000,
	})
	if tableBz == nil {
		return
	}
	_ = tableBz.MergeCellsVertical(1, 7, 0)
	tableBz.MergeCellsVertical(9, 10, 0)
	tableBz.MergeCellsVertical(12, 13, 0)
	tableBz.MergeCellsVertical(14, 15, 0)
	
	tableBz.SetCellText(0, 0, "适用范围")
	tableBz.SetCellText(0, 1, "文件名")
	tableBz.SetCellText(0, 2, "文件编号")
	tableBz.SetCellText(1, 0, "生态环境部")
	tableBz.SetCellText(1, 1, "《石化行业VOCs污染源排查工作指南》")
	
	tableBz.SetCellText(2, 1, "《泄漏和敞开液面排放的挥发性有机物检测技术导则》")
	tableBz.SetCellText(2, 2, "HJ 733-2014")
	tableBz.SetCellText(3, 1, "《工业企业挥发性有机物泄漏检测与修复技术指南》")
	tableBz.SetCellText(3, 2, "HJ 1230-2021")
	tableBz.SetCellText(4, 1, "《石油炼制工业污染源排放标准》")
	tableBz.SetCellText(4, 2, "GB 31570-2015")
	tableBz.SetCellText(5, 1, "《石油化学工业污染物排放标准》")
	tableBz.SetCellText(5, 2, "GB 31571-2015")
	tableBz.SetCellText(6, 1, "《挥发性有机物无组织排放控制标准》")
	tableBz.SetCellText(6, 2, "GB 37822-2019")
	tableBz.SetCellText(7, 1, "《制药工业大气污染物排放标准》")
	tableBz.SetCellText(7, 2, "GB 37823-2019")
	tableBz.SetCellText(8, 0, "南京")
	tableBz.SetCellText(8, 1, "《设备与管线组件挥发性有机物泄漏控制技术规范》")
	tableBz.SetCellText(8, 2, "DB3201/T1228-2024")
	tableBz.SetCellText(9, 0, "江苏南京园区")
	tableBz.SetCellText(9, 1, "《南京化工园区企业挥发性气体无泄漏检测规程》及《南京化工园区在线设备选型指南》的通知")
	tableBz.SetCellText(9, 2, "宁化环字〔2015〕38号")
	tableBz.SetCellText(10, 1, "《南京江北新材料科技园化工企业大修期间环境管控方案》的通知")
	tableBz.SetCellText(10, 2, "宁新区新科办发〔2020〕60号")
	tableBz.SetCellText(11, 0, "长江三角洲")
	tableBz.SetCellText(11, 1, "《设备泄漏挥发性有机物排放控制技术规范》")
	tableBz.SetCellText(11, 2, "DB34/T310007-2021")
	tableBz.SetCellText(12, 0, "广东")
	tableBz.SetCellText(12, 1, "《广东省泄漏检测与修复（LDAR）实施技术规范》")
	tableBz.SetCellText(12, 2, "粤环函〔2016〕1049号")
	tableBz.SetCellText(13, 1, "《广东省泄漏检测与维修制度（LDAR）实施技术要求》")
	tableBz.SetCellText(13, 2, "粤环函〔2013〕830号")
	tableBz.SetCellText(14, 0, "天津")
	tableBz.SetCellText(14, 1, "《天津市工业企业挥发性有机物排放控制标准》")
	tableBz.SetCellText(14, 2, "DB12-524-2014")
	tableBz.SetCellText(15, 1, "天津《工业企业挥发性有机物排放控制标准》")
	tableBz.SetCellText(15, 2, "DB12/524-2020")
	tableBz.SetCellText(16, 0, "河北")
	tableBz.SetCellText(16, 1, "工业企业挥发性有机物排放控制标准")
	tableBz.SetCellText(16, 2, "DB 132322-2016")
	tableBz.SetCellText(17, 0, "山东")
	tableBz.SetCellText(17, 1, "石油炼制工业泄漏检测与修复实施技术要求")
	tableBz.SetCellText(17, 2, "DB 37—2016")
	tableBz.SetCellText(18, 0, "使用范围")
	tableBz.SetCellText(18, 1, "文件名")
	tableBz.SetCellText(18, 2, "文件编号")
	tableBz.SetCellText(19, 0, "四川")
	tableBz.SetCellText(19, 1, "四川省挥发性有机物泄漏检测与修复（LDAR）实施技术规范")
	tableBz.SetCellText(19, 2, "/")
	tableBz.SetCellText(20, 0, "河南")
	tableBz.SetCellText(20, 1, "工业企业挥发性有机物泄漏检测与修复技术规范")
	tableBz.SetCellText(20, 2, "DB 41T2364-2022")
	tableBz.SetCellText(21, 0, "浙江嘉兴")
	tableBz.SetCellText(21, 1, "《嘉兴港区泄漏检测与修复体系（LDAR） 建设管理办法》")
	tableBz.SetCellText(21, 2, "/")
	tableBz.SetCellText(22, 0, "新疆")
	tableBz.SetCellText(22, 1, "《新疆维吾尔族自治区工业企业挥发性有机物泄漏检测与修复（LDAR）技术要求试行》")
	tableBz.SetCellText(22, 2, "/")
	tableBz.SetCellText(22, 0, "宁夏石嘴山")
	tableBz.SetCellText(22, 1, "《石嘴山市环境保护局（关于在化工企业开展泄漏检测与修复）》")
	tableBz.SetCellText(22, 2, "石环通字〔2018〕46号")
	
	for i := 0; i < tableBz.GetRowCount(); i++ {
		tableBz.SetRowHeight(i, &RowHeightConfig{
			Height: 33,
			Rule:   RowHeightMinimum,
		})
	}
	
	// 添加分节符，切换回竖版（保持页码连续，继承页眉页脚）
	sectionBreak3 := doc.AddParagraph("")
	sectionBreak3.AddSectionBreakWithStartPage(OrientationPortrait, doc, 0, true)
	
	// 添加更多标题用于测试目录显示
	// 一级标题：四号（14磅），宋体，加粗
	bookmarkName6 := "_Toc_第七页标题"
	p6 := doc.AddHeadingParagraphWithBookmark("第七页标题", 1, bookmarkName6)
	if len(p6.Runs) > 0 {
		if p6.Runs[0].Properties == nil {
			p6.Runs[0].Properties = &RunProperties{}
		}
		// 设置字体格式：四号（14磅 * 2 = 28）
		p6.Runs[0].Properties.FontFamily = &FontFamily{ASCII: "SimSun", HAnsi: "SimSun", EastAsia: "SimSun"}
		p6.Runs[0].Properties.FontSize = &FontSize{Val: "28"}
		p6.Runs[0].Properties.Color = &Color{Val: "000000"}
		p6.Runs[0].Properties.Bold = &Bold{}
	}
	
	// 添加二级标题：小四（12磅），宋体，加粗
	bookmarkName6_1 := "_Toc_第七页二级标题1"
	p6_1 := doc.AddHeadingParagraphWithBookmark("第七页二级标题1", 2, bookmarkName6_1)
	if len(p6_1.Runs) > 0 {
		if p6_1.Runs[0].Properties == nil {
			p6_1.Runs[0].Properties = &RunProperties{}
		}
		// 设置字体格式：小四（12磅 * 2 = 24）
		p6_1.Runs[0].Properties.FontFamily = &FontFamily{ASCII: "SimSun", HAnsi: "SimSun", EastAsia: "SimSun"}
		p6_1.Runs[0].Properties.FontSize = &FontSize{Val: "24"}
		p6_1.Runs[0].Properties.Color = &Color{Val: "000000"}
		p6_1.Runs[0].Properties.Bold = &Bold{}
	}
	
	// 添加三级标题：小四（12磅），宋体，加粗
	bookmarkName6_1_1 := "_Toc_第七页三级标题1"
	p6_1_1 := doc.AddHeadingParagraphWithBookmark("第七页三级标题1", 3, bookmarkName6_1_1)
	if len(p6_1_1.Runs) > 0 {
		if p6_1_1.Runs[0].Properties == nil {
			p6_1_1.Runs[0].Properties = &RunProperties{}
		}
		// 设置字体格式：小四（12磅 * 2 = 24）
		p6_1_1.Runs[0].Properties.FontFamily = &FontFamily{ASCII: "SimSun", HAnsi: "SimSun", EastAsia: "SimSun"}
		p6_1_1.Runs[0].Properties.FontSize = &FontSize{Val: "24"}
		p6_1_1.Runs[0].Properties.Color = &Color{Val: "000000"}
		p6_1_1.Runs[0].Properties.Bold = &Bold{}
	}
	
	// 添加另一个三级标题：小四（12磅），宋体，加粗
	bookmarkName6_1_2 := "_Toc_第七页三级标题2"
	p6_1_2 := doc.AddHeadingParagraphWithBookmark("第七页三级标题2", 3, bookmarkName6_1_2)
	if len(p6_1_2.Runs) > 0 {
		if p6_1_2.Runs[0].Properties == nil {
			p6_1_2.Runs[0].Properties = &RunProperties{}
		}
		// 设置字体格式：小四（12磅 * 2 = 24）
		p6_1_2.Runs[0].Properties.FontFamily = &FontFamily{ASCII: "SimSun", HAnsi: "SimSun", EastAsia: "SimSun"}
		p6_1_2.Runs[0].Properties.FontSize = &FontSize{Val: "24"}
		p6_1_2.Runs[0].Properties.Color = &Color{Val: "000000"}
		p6_1_2.Runs[0].Properties.Bold = &Bold{}
	}
	
	// 添加另一个二级标题：小四（12磅），宋体，加粗
	bookmarkName6_2 := "_Toc_第七页二级标题2"
	p6_2 := doc.AddHeadingParagraphWithBookmark("第七页二级标题2", 2, bookmarkName6_2)
	if len(p6_2.Runs) > 0 {
		if p6_2.Runs[0].Properties == nil {
			p6_2.Runs[0].Properties = &RunProperties{}
		}
		// 设置字体格式：小四（12磅 * 2 = 24）
		p6_2.Runs[0].Properties.FontFamily = &FontFamily{ASCII: "SimSun", HAnsi: "SimSun", EastAsia: "SimSun"}
		p6_2.Runs[0].Properties.FontSize = &FontSize{Val: "24"}
		p6_2.Runs[0].Properties.Color = &Color{Val: "000000"}
		p6_2.Runs[0].Properties.Bold = &Bold{}
	}
	
	// 添加内容段落
	contentPara6 := doc.AddParagraph("这是第七页的内容，用于测试多级标题在目录中的显示效果。")
	if len(contentPara6.Runs) > 0 {
		if contentPara6.Runs[0].Properties == nil {
			contentPara6.Runs[0].Properties = &RunProperties{}
		}
		contentPara6.Runs[0].Properties.FontFamily = &FontFamily{ASCII: "SimSun", HAnsi: "SimSun", EastAsia: "SimSun"}
		contentPara6.Runs[0].Properties.FontSize = &FontSize{Val: "24"}
		contentPara6.Runs[0].Properties.Color = &Color{Val: "000000"}
	}
	
	// 在目录页位置生成目录
	tocConfig := &TOCConfig{
		Title:        "目录",
		MaxLevel:     3,
		ShowPageNum:  true,
		RightAlign:   true,
		UseHyperlink: true,
		DotLeader:    true,
		PageOffset:   0, // 过滤掉封面、目录、第1-4页（共6页）
	}
	
	// 调用toc.go中的方法生成目录
	if err := doc.GenerateTOCAtPosition(tocConfig, tocInsertIndex, tocInsertIndex); err != nil {
		t.Error(err)
	}
	
	// 保存文档
	outputPath := "test_fixed.docx"
	if err := doc.Save(outputPath); err != nil {
		t.Error(err)
		return
	}
	
	t.Logf("文档已保存: %s", outputPath)
	t.Logf("预期效果：")
	t.Logf("  - 封面、目录、第1-4页不显示页码")
	t.Logf("  - 从第5页开始显示页码，页码从1开始")
	t.Logf("  - 竖版和横版页面交替，页码连续：1, 2, 3, ...")
	t.Logf("提示：打开Word文档后，按Ctrl+A全选，然后按F9更新所有字段，即可更新目录页码")
}
