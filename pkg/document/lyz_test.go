// @Author: liyongzhen
// @Description:
// @File: lyz_test
// @Date: 2025/6/19 16:10

package document

import (
	"fmt"
	"testing"
	"time"

	"github.com/ZeroHawkeye/wordZero/pkg/style"
)

func TestHeaderStyle(t *testing.T) {
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
	//coverContent.SetFontSize(14)
	textFormat := &TextFormat{
		FontFamily: "SimSun",
		FontSize:   14,
		FontColor:  "000000",
		Bold:       true,
	}
	reportExplain(doc, textFormat, &SpacingConfig{})
	reportExplain(doc, textFormat, &SpacingConfig{})

	doc.AddParagraph("").AddPageBreak()

	// 目录页 - 先创建目录标题占位符，后续会被目录内容替换
	tocTitlePara := doc.AddParagraph("目录")
	tocTitlePara.SetStyle(style.StyleHeading1)

	// 记录目录插入位置（占位符段落的位置）
	tocInsertIndex := len(doc.Body.Elements) - 1

	//doc.AddParagraph("").AddPageBreak()
	// 目录之后，应该是第1页
	shuban(doc, textFormat)

	// 添加分节符，从这里开始启用新的节
	sectionBreak := doc.AddParagraph("")
	sectionBreak.AddSectionBreak(OrientationPortrait, doc)

	// 在新节上配置页眉页脚（只影响新节及之后内容）
	// 注意：必须在分节符之后设置页眉页脚
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
	// 重置页码从1开始（在设置页眉页脚之后）
	//doc.RestartPageNumber()

	// 使用 AddHeadingParagraphWithBookmark 创建标题，这样可以被目录生成功能识别并支持跳转
	bookmarkName1 := "_Toc_第二页标题"
	p := doc.AddHeadingParagraphWithBookmark("第二页标题", 1, bookmarkName1)
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

	// 在标题后添加一些内容，确保页面有内容
	contentPara := doc.AddParagraph("这是第二页标题页面的内容。")
	if len(contentPara.Runs) > 0 {
		if contentPara.Runs[0].Properties == nil {
			contentPara.Runs[0].Properties = &RunProperties{}
		}
		contentPara.Runs[0].Properties.FontFamily = &FontFamily{ASCII: "SimSun", HAnsi: "SimSun", EastAsia: "SimSun"}
		contentPara.Runs[0].Properties.FontSize = &FontSize{Val: "24"} // 12磅 * 2
		contentPara.Runs[0].Properties.Color = &Color{Val: "000000"}
	}

	// 在内容段落后添加分页符和分节符
	contentPara.AddPageBreak()
	contentPara.SetSpacing(&SpacingConfig{
		BeforePara: 0,
		AfterPara:  0,
	})

	contentPara.AddSectionBreak(OrientationLandscape, doc)

	// 标题段落也需要设置间距
	p.SetSpacing(&SpacingConfig{
		BeforePara: 0,
		AfterPara:  0,
	})

	// 在横版节中再次设置页眉页脚（可选，如果需要不同页眉页脚）
	// 注意：如果要保持页码连续，不要再次调用RestartPageNumber()
	err = doc.AddStyleHeader(HeaderFooterTypeDefault, "xxx科技有限公司\nRLHB", "2025010", &TextFormat{
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
	// 不要在这里再次重置页码，保持页码连续

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

	// 使用 AddHeadingParagraphWithBookmark 创建标题，这样可以被目录生成功能识别并支持跳转
	textFormat.Bold = true
	textFormat.FontSize = 14
	shuban(doc, textFormat)

	// 在目录页位置生成目录
	tocConfig := &TOCConfig{
		Title:        "目录",
		MaxLevel:     3,
		ShowPageNum:  true,
		RightAlign:   true,
		UseHyperlink: true,
		DotLeader:    true,
	}

	// 调用toc.go中的方法生成目录
	if err := doc.GenerateTOCAtPosition(tocConfig, tocInsertIndex, tocInsertIndex); err != nil {
		t.Error(err)
	}

	//doc.UpdateTOC()

	// 保存文档
	outputPath := "test.docx"
	if err := doc.Save(outputPath); err != nil {
		t.Error(err)
		return
	}

	// 注意：目录页码使用PAGEREF字段，初始值为占位符1
	// 打开Word文档后，按Ctrl+A全选，然后按F9更新所有字段，即可更新目录页码
	t.Logf("文档已保存: %s", outputPath)
	t.Logf("提示：打开Word文档后，按Ctrl+A全选，然后按F9更新所有字段，即可更新目录页码")
}

func shuban(doc *Document, textFormat *TextFormat) {
	bookmarkName2 := "_Toc_3_企业基本信息"
	p1 := doc.AddHeadingParagraphWithBookmark("3 企业基本信息", 1, bookmarkName2)
	// 如果需要自定义格式，可以更新Run的属性
	if len(p1.Runs) > 0 {
		if p1.Runs[0].Properties == nil {
			p1.Runs[0].Properties = &RunProperties{}
		}
		// 设置字体格式
		p1.Runs[0].Properties.FontFamily = &FontFamily{ASCII: "SimSun", HAnsi: "SimSun", EastAsia: "SimSun"}
		p1.Runs[0].Properties.FontSize = &FontSize{Val: "28"} // 14磅 * 2
		p1.Runs[0].Properties.Color = &Color{Val: "000000"}
		p1.Runs[0].Properties.Bold = &Bold{}
	}
	p1.AddPageBreak()

	// 如果需要在此节中设置页眉页脚（保持与之前相同的页眉页脚）
	// 由于前面已经设置了页眉页脚，这里无需重复设置
	// 但需要确保当前节的页码继续增加

	textFormat.Bold = false
	textFormat.FontSize = 12
	doc.AddFormattedParagraph("我的来急啦圣诞节啦解放啦解放啦是老大解放啦卡随机发", textFormat)

	textFormat.FontSize = 16
	content := doc.AddFormattedParagraph("委托单位：", textFormat)
	content.SetSpacing(&SpacingConfig{
		FirstLineIndent: 5 * 22, // 5个缩进符
		LineSpacing:     1.5,
	})
	content.AddRun("上海中科", textFormat, &RunProperties{
		Underline: &Underline{
			Val: "single",
		},
	})
	content = doc.AddFormattedParagraph("编制单位：", textFormat)
	content.SetSpacing(&SpacingConfig{
		FirstLineIndent: 5 * 22, // 5个缩进符
		LineSpacing:     1.5,
	})
	content.AddRun("河南瑞蓝环保科技有限公司", textFormat, &RunProperties{
		Underline: &Underline{
			Val: "single",
		},
	})
	content = doc.AddFormattedParagraph("编制日期：", textFormat)
	content.SetSpacing(&SpacingConfig{
		FirstLineIndent: 5 * 22, // 5个缩进符
		LineSpacing:     1.5,
	})
	content.AddRun(fmt.Sprintf("%s  年  %s  月  %s  日", time.Now().Format("2006"), time.Now().Format("01"), time.Now().Format("02")), textFormat, &RunProperties{
		Underline: &Underline{
			Val: "single",
		},
	})

	// 下标测试
	content = doc.AddFormattedParagraph("x", textFormat)
	content.AddRun("1", textFormat, &RunProperties{
		VertAlign: &VertAlign{Val: "subscript"}, // ⬅️ 添加下标
	})
	content.AddRun("slss", textFormat, &RunProperties{})
}

func reportExplain(doc *Document, textFormat *TextFormat, spacingConfig *SpacingConfig) {
	doc.AddParagraph("")
	textFormat.FontSize = 24
	content := doc.AddFormattedParagraph("检测报告说明", textFormat)
	content.AddPageBreak()
	content.SetStyle(style.StyleNormal)
	content.SetAlignment(AlignCenter)
	spacingConfig.BeforePara = 0
	spacingConfig.AfterPara = 30
	content.SetSpacing(spacingConfig)

	textFormat.FontSize = 15
	//doc.AddImageFromFile(picDir+"zhang.png",)
	content = doc.AddFormattedParagraph("1、本公司检测报告须同时具有检验检测专用章、骑缝章及CMA章标志，缺少其中之一则报告无效。", textFormat)
	content.SetStyle(style.StyleNormal)
	content.SetAlignment(AlignLeft)
	spacingConfig.AfterPara = 0
	spacingConfig.LineSpacing = 1.5
	spacingConfig.FirstLineIndent = 22
	content.SetSpacing(spacingConfig)

	content = doc.AddFormattedParagraph("2、结果表述清晰，易于理解。无授权签字人签字识别的，报告无效。", textFormat)
	content.SetStyle(style.StyleNormal)
	content.SetAlignment(AlignLeft)
	content.SetSpacing(spacingConfig)

	content = doc.AddFormattedParagraph("3、当需要对检测报告做出意见和解释时，本公司依据评审准则将意见和解释在报告中清晰标注。", textFormat)
	content.SetStyle(style.StyleNormal)
	content.SetAlignment(AlignLeft)
	content.SetSpacing(spacingConfig)

	content = doc.AddFormattedParagraph("4、本报告未经同意不得用于广告宣传，复制本报告中的部分内容无效。", textFormat)
	content.SetStyle(style.StyleNormal)
	content.SetAlignment(AlignLeft)
	spacingConfig.BeforePara = 0
	spacingConfig.AfterPara = 130
	content.SetSpacing(spacingConfig)

	content = doc.AddFormattedParagraph("编制单位：河南瑞蓝环保科技有限公司", textFormat)
	content.SetStyle(style.StyleNormal)
	content.SetAlignment(AlignLeft)
	spacingConfig.AfterPara = 0
	content.SetSpacing(spacingConfig)

	content = doc.AddFormattedParagraph("单位地址：河南省郑州市高新技术产业开发区西三环路283号11号楼5层30号", textFormat)
	content.SetStyle(style.StyleNormal)
	content.SetAlignment(AlignLeft)
	content.SetSpacing(spacingConfig)

	content = doc.AddFormattedParagraph("联系电话：0371-86557168", textFormat)
	content.SetStyle(style.StyleNormal)
	content.SetAlignment(AlignLeft)
	content.SetSpacing(spacingConfig)

	content = doc.AddFormattedParagraph("邮    箱：hnrlhbkj@126.com", textFormat)
	content.SetStyle(style.StyleNormal)
	content.SetAlignment(AlignLeft)
	content.SetSpacing(spacingConfig)
}
