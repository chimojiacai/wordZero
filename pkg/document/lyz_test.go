// @Author: liyongzhen
// @Description:
// @File: lyz_test
// @Date: 2025/6/19 16:10

package document

import "testing"

func TestHeaderStyle(t *testing.T) {
	doc := New()
	p := doc.AddParagraph("heiheihei!!!")
	p.AddLineBreak("第二行")

	p.AddPageBreak()
	err := doc.AddStyleHeader(HeaderFooterTypeDefault, "xxx科技有限公司\nRLHB", "2025010", &TextFormat{
		FontFamily: "SimSun",
		FontSize:   9,
		FontColor:  "000000",
	})
	if err != nil {
		t.Error(err)
	}
	//pH.SetAlignment(AlignCenter)
	doc.AddParagraph("heiheihei!!!2")
	doc.Save("test.docx")
}
