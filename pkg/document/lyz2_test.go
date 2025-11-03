// @Author: liyongzhen
// @Description:
// @File: lyz2_test
// @Date: 2025/11/3 18:01

package document

import (
	"fmt"
	"os"
	"testing"

	"github.com/ZeroHawkeye/wordZero/pkg/style"
)

func splitLongTextToLines(text string, maxLen int) []string {
	var lines []string
	for i := 0; i < len(text); i += maxLen {
		end := i + maxLen
		if end > len(text) {
			end = len(text)
		}
		lines = append(lines, text[i:end])
	}
	return lines
}
func TestWordZero(t *testing.T) {
	doc := New()
	textFormat := &TextFormat{}
	textFormat.Bold = true
	textFormat.FontSize = 14
	content := doc.AddFormattedParagraph("附件11 现场工作照", textFormat)
	content.AddPageBreak() // 分页
	content.SetStyle(style.StyleHeading2)
	content.SetSpacing(&SpacingConfig{
		BeforePara:  1,
		LineSpacing: 1.5,
	})
	picDirs := []string{"img.png", "img_1.png", "img_2.png", "img_3.png"}
	imgInfos := make([]*ImageInfo, 0)
	for _, v := range picDirs {
		data, _ := os.ReadFile(v)
		img, err := doc.AddImageFromDataWithoutElement(
			data,
			v,              // fileName e.g. "img_1.png"
			ImageFormatPNG, // or ImageFormatPNG
			70, 93,         // width/height in pt
			&ImageConfig{
				Size:      &ImageSize{Width: 70.4, Height: 93.9},
				Position:  ImagePositionInline,
				Alignment: AlignCenter,
			},
		)
		if err != nil {
			t.Fatalf("图片注册失败: %s", v)
		}
		imgInfos = append(imgInfos, img)
	}
	if err := doc.InsertImageRow(imgInfos, ""); err != nil {
		t.Fatalf("插入图片失败: %v", err)
	}

	// 正确设置最后一页为横向的方法
	// 我们需要将分节符添加到文档主体的末尾，而不是段落中
	sectPr := &SectionProperties{
		PageSize: &PageSizeXML{
			W:      "16838", // A4 横向宽度
			H:      "11906", // A4 横向高度
			Orient: "landscape",
		},
		PageMargins: &PageMargin{
			Top:    fmt.Sprintf("%.0f", mmToTwips(doc.GetPageSettings().MarginTop)),
			Bottom: fmt.Sprintf("%.0f", mmToTwips(doc.GetPageSettings().MarginBottom)),
			Left:   fmt.Sprintf("%.0f", mmToTwips(doc.GetPageSettings().MarginLeft)),
			Right:  fmt.Sprintf("%.0f", mmToTwips(doc.GetPageSettings().MarginRight)),
		},
	}

	// 将节属性直接添加到文档主体末尾
	doc.Body.Elements = append(doc.Body.Elements, sectPr)

	doc.Save("附件11现场工作照.docx")
}
