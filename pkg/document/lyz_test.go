// @Author: liyongzhen
// @Description:
// @File: lyz_test
// @Date: 2025/6/19 16:10

package document

import (
	"github.com/ZeroHawkeye/wordZero/pkg/style"
	"testing"
)

func TestHeaderStyle(t *testing.T) {
	doc := New()
	err := doc.AddStyleHeader(HeaderFooterTypeDefault, "xxx科技有限公司\nRLHB", "2025010", &TextFormat{
		FontFamily: "SimSun",
		FontSize:   9,
		FontColor:  "000000",
	})
	if err != nil {
		t.Error(err)
	}

	textFormat := &TextFormat{
		FontFamily: "SimSun",
		FontSize:   14,
		FontColor:  "000000",
		Bold:       true,
	}
	doc.AddFormattedParagraph("2 标准依据", textFormat).SetStyle(style.StyleHeading1)

	//textFormat.Bold = false
	//textFormat.FontSize = 12
	//// 表格:标准依据
	//tableBz := doc.AddTable(&TableConfig{
	//	Rows:      23,
	//	Cols:      3,
	//	ColWidths: []int{2000, 4000, 3000},
	//	Width:     9000,
	//})
	//if tableBz == nil {
	//	return
	//}
	//_ = tableBz.MergeCellsVertical(1, 7, 0)
	//tableBz.MergeCellsVertical(9, 10, 0)
	//tableBz.MergeCellsVertical(12, 13, 0)
	//tableBz.MergeCellsVertical(14, 15, 0)
	//
	//tableBz.SetCellText(0, 0, "适用范围")
	//tableBz.SetCellText(0, 1, "文件名")
	//tableBz.SetCellText(0, 2, "文件编号")
	//tableBz.SetCellText(1, 0, "生态环境部")
	//tableBz.SetCellText(1, 1, "《石化行业VOCs污染源排查工作指南》")
	//tableBz.SetCellText(1, 2, "环办〔2015〕104号")
	//tableBz.SetCellText(2, 1, "《泄漏和敞开液面排放的挥发性有机物检测技术导则》")
	//tableBz.SetCellText(2, 2, "HJ 733-2014")
	//tableBz.SetCellText(3, 1, "《工业企业挥发性有机物泄漏检测与修复技术指南》")
	//tableBz.SetCellText(3, 2, "HJ 1230-2021")
	//tableBz.SetCellText(4, 1, "《石油炼制工业污染源排放标准》")
	//tableBz.SetCellText(4, 2, "GB 31570-2015")
	//tableBz.SetCellText(5, 1, "《石油化学工业污染物排放标准》")
	//tableBz.SetCellText(5, 2, "GB 31571-2015")
	//tableBz.SetCellText(6, 1, "《挥发性有机物无组织排放控制标准》")
	//tableBz.SetCellText(6, 2, "GB 37822-2019")
	//tableBz.SetCellText(7, 1, "《制药工业大气污染物排放标准》")
	//tableBz.SetCellText(7, 2, "GB 37823-2019")
	//tableBz.SetCellText(8, 0, "南京")
	//tableBz.SetCellText(8, 1, "《设备与管线组件挥发性有机物泄漏控制技术规范》")
	//tableBz.SetCellText(8, 2, "DB3201/T1228-2024")
	//tableBz.SetCellText(9, 0, "江苏南京园区")
	//tableBz.SetCellText(9, 1, "《南京化工园区企业挥发性气体无泄漏检测规程》及《南京化工园区在线设备选型指南》的通知")
	//tableBz.SetCellText(9, 2, "宁化环字〔2015〕38号")
	//tableBz.SetCellText(10, 1, "《南京江北新材料科技园化工企业大修期间环境管控方案》的通知")
	//tableBz.SetCellText(10, 2, "宁新区新科办发〔2020〕60号")
	//tableBz.SetCellText(11, 0, "长江三角洲")
	//tableBz.SetCellText(11, 1, "《设备泄漏挥发性有机物排放控制技术规范》")
	//tableBz.SetCellText(11, 2, "DB34/T310007-2021")
	//tableBz.SetCellText(12, 0, "广东")
	//tableBz.SetCellText(12, 1, "《广东省泄漏检测与修复（LDAR）实施技术规范》")
	//tableBz.SetCellText(12, 2, "粤环函〔2016〕1049号")
	//tableBz.SetCellText(13, 1, "《广东省泄漏检测与维修制度（LDAR）实施技术要求》")
	//tableBz.SetCellText(13, 2, "粤环函〔2013〕830号")
	//tableBz.SetCellText(14, 0, "天津")
	//tableBz.SetCellText(14, 1, "《天津市工业企业挥发性有机物排放控制标准》")
	//tableBz.SetCellText(14, 2, "DB12-524-2014")
	//tableBz.SetCellText(15, 1, "天津《工业企业挥发性有机物排放控制标准》")
	//tableBz.SetCellText(15, 2, "DB12/524-2020")
	//tableBz.SetCellText(16, 0, "河北")
	//tableBz.SetCellText(16, 1, "工业企业挥发性有机物排放控制标准")
	//tableBz.SetCellText(16, 2, "DB 132322-2016")
	//tableBz.SetCellText(17, 0, "山东")
	//tableBz.SetCellText(17, 1, "石油炼制工业泄漏检测与修复实施技术要求")
	//tableBz.SetCellText(17, 2, "DB 37—2016")
	//tableBz.SetCellText(18, 0, "使用范围")
	//tableBz.SetCellText(18, 1, "文件名")
	//tableBz.SetCellText(18, 2, "文件编号")
	//tableBz.SetCellText(19, 0, "四川")
	//tableBz.SetCellText(19, 1, "四川省挥发性有机物泄漏检测与修复（LDAR）实施技术规范")
	//tableBz.SetCellText(19, 2, "/")
	//tableBz.SetCellText(20, 0, "河南")
	//tableBz.SetCellText(20, 1, "工业企业挥发性有机物泄漏检测与修复技术规范")
	//tableBz.SetCellText(20, 2, "DB 41T2364-2022")
	//tableBz.SetCellText(21, 0, "浙江嘉兴")
	//tableBz.SetCellText(21, 1, "《嘉兴港区泄漏检测与修复体系（LDAR） 建设管理办法》")
	//tableBz.SetCellText(21, 2, "/")
	//tableBz.SetCellText(22, 0, "新疆")
	//tableBz.SetCellText(22, 1, "《新疆维吾尔族自治区工业企业挥发性有机物泄漏检测与修复（LDAR）技术要求试行》")
	//tableBz.SetCellText(22, 2, "/")
	//tableBz.SetCellText(22, 0, "宁夏石嘴山")
	//tableBz.SetCellText(22, 1, "《石嘴山市环境保护局（关于在化工企业开展泄漏检测与修复）》")
	//tableBz.SetCellText(22, 2, "石环通字〔2018〕46号")
	//
	//for i := 0; i < tableBz.GetRowCount(); i++ {
	//	tableBz.SetRowHeight(i, &RowHeightConfig{
	//		Height: 33,
	//		Rule:   RowHeightMinimum,
	//	})
	//}

	doc.AddParagraph("").AddPageBreak() // 添加分页

	doc.AddParagraph("第二页第一行") // 正确位置
	doc.AddParagraph("第二页第2行") // 正确位置

	textFormat.Bold = true
	textFormat.FontSize = 14
	doc.AddFormattedParagraph("3 企业基本信息", textFormat).SetStyle(style.StyleHeading1)
	textFormat.Bold = false
	textFormat.FontSize = 12
	doc.AddFormattedParagraph("我的来急啦圣诞节啦解放啦解放啦是老大解放啦卡随机发", textFormat)
	doc.Save("test.docx")
}
