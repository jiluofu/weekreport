package main

import (

    "fmt"
    "io/ioutil"
    "time"

    "github.com/bitly/go-simplejson"
    "github.com/xuri/excelize"
)


func main() {

    row := 1
    dat, _ := ioutil.ReadFile("/Users/zhuxu/Documents/weekreport/test.json")
    // fmt.Println(string(dat))

    json, err := simplejson.NewJson(dat)
    if err != nil {

        fmt.Println("error")
    }

    t := time.Now()
    const day_secs = 24 * 3600

    weekday_today := int64(t.Weekday())
    if (weekday_today > 5) {

        weekday_today = 5
    }

    weekday1 := time.Unix(t.Unix() - day_secs * (weekday_today - 1), 0)
    weekday5 := time.Unix(t.Unix() + day_secs * (5 - weekday_today), 0)
    day1 := weekday1.Format("01.02")
    day5 := weekday5.Format("01.02")
    fmt.Println(day1)
    fmt.Println(day5)

    xlsx := excelize.NewFile()

    style, err := xlsx.NewStyle(`{"alignment":{"horizontal":"top","ident":1,"justify_last_line":true,"reading_order":0,"relative_indent":1,"vertical":"top","wrap_text":false},"border":[{"type":"left","color":"DDDDDD","style":1}, {"type":"right","color":"DDDDDD","style":1}, {"type":"top","color":"DDDDDD","style":1}, {"type":"bottom","color":"DDDDDD","style":1}]}`)
    if err != nil {
        fmt.Println(err)
    }
    xlsx.SetCellStyle("Sheet1", "A1", "H50", style)

    style_boldfill, _ := xlsx.NewStyle(`{"font":{"bold":true,"italic":false},"fill":{"type":"pattern","color":["#F0F0F0"], "pattern":1},"border":[{"type":"left","color":"DDDDDD","style":1}, {"type":"right","color":"DDDDDD","style":1}, {"type":"top","color":"DDDDDD","style":1}, {"type":"bottom","color":"DDDDDD","style":1}]}`)
    style_bold, _ := xlsx.NewStyle(`{"font":{"bold":true,"italic":false},"alignment":{"horizontal":"top","ident":1,"justify_last_line":true,"reading_order":0,"relative_indent":1,"vertical":"top","wrap_text":false},"border":[{"type":"left","color":"DDDDDD","style":1}, {"type":"right","color":"DDDDDD","style":1}, {"type":"top","color":"DDDDDD","style":1}, {"type":"bottom","color":"DDDDDD","style":1}]}`)
    style_border, _ := xlsx.NewStyle(`{"font":{"bold":false,"italic":false},"alignment":{"horizontal":"top","ident":1,"justify_last_line":true,"reading_order":0,"relative_indent":1,"vertical":"top","wrap_text":false},"border":[{"type":"left","color":"DDDDDD","style":1}, {"type":"right","color":"DDDDDD","style":1}, {"type":"top","color":"DDDDDD","style":1}, {"type":"bottom","color":"DDDDDD","style":1}]}`)

    // Create a new sheet.
    xlsx.NewSheet(1, "Sheet1")
    xlsx.SetCellValue("Sheet1", fmt.Sprintf("A%d", row), fmt.Sprintf("个人周报[%s-%s]", day1, day5))
    xlsx.MergeCell("sheet1", fmt.Sprintf("A%d", row), fmt.Sprintf("H%d", row))
    xlsx.SetCellStyle("Sheet1", "A1", "A1", style_boldfill)
    row ++

    content := "本周总结：\n"
    this_week, _ := json.Get("this_week").Array()
    for i := 0; i < len(this_week); i++ {

        content += fmt.Sprintf("%d、%s\n", i + 1, this_week[i].(string))
    }

    content += "\n下周计划：\n"

    next_week, _ := json.Get("next_week").Array()
    for i := 0; i < len(next_week); i ++ {

        // fmt.Println(next_week[i])
        content += fmt.Sprintf("%d、%s\n", i + 1, next_week[i].(string))
    }
    fmt.Println(content)
    xlsx.SetCellStyle("Sheet1", "A2", "A2", style_border)
    xlsx.SetCellValue("Sheet1", fmt.Sprintf("A%d", row), content)
    xlsx.MergeCell("sheet1", fmt.Sprintf("A%d", row), fmt.Sprintf("H%d", row))
    row ++

    xlsx.SetCellValue("Sheet1", fmt.Sprintf("A%d", row), "项目")
    xlsx.SetCellValue("Sheet1", fmt.Sprintf("B%d", row), "项目内容")
    xlsx.SetCellValue("Sheet1", fmt.Sprintf("C%d", row), "人员")
    xlsx.SetCellValue("Sheet1", fmt.Sprintf("D%d", row), "分解")
    xlsx.SetCellValue("Sheet1", fmt.Sprintf("E%d", row), "计划时间")
    xlsx.SetCellValue("Sheet1", fmt.Sprintf("F%d", row), "")
    xlsx.SetCellValue("Sheet1", fmt.Sprintf("G%d", row), "状态")
    xlsx.SetCellValue("Sheet1", fmt.Sprintf("H%d", row), "备注")
    xlsx.MergeCell("sheet1", fmt.Sprintf("E%d", row), fmt.Sprintf("F%d", row))
    xlsx.SetCellStyle("Sheet1", fmt.Sprintf("A%d", row), fmt.Sprintf("H%d", row), style_bold)
    row ++

    projects, _ := json.Get("projects").Array()

    for i := 0; i < len(projects); i ++ {

        project, _ := projects[i].(map[string]interface{})

        xlsx.SetCellValue("Sheet1", fmt.Sprintf("A%d", row), project["title"])
        fmt.Println(project["content"])
        xlsx.SetCellValue("Sheet1", fmt.Sprintf("B%d", row), fmt.Sprintf("%s", project["content"]))
        xlsx.SetCellValue("Sheet1", fmt.Sprintf("C%d", row), project["staffs"])
        xlsx.SetCellStyle("Sheet1", fmt.Sprintf("A%d", row), fmt.Sprintf("A%d", row), style_bold)

        details, _ := project["details"].([]interface{})
        for j := 0; j < len(details); j ++ {

            detail, _ := details[j].(map[string]interface{})

            xlsx.SetCellValue("Sheet1", fmt.Sprintf("D%d", row + j), detail["title"])
            xlsx.SetCellValue("Sheet1", fmt.Sprintf("E%d", row + j), detail["start_time"])
            xlsx.SetCellValue("Sheet1", fmt.Sprintf("F%d", row + j), detail["end_time"])
            xlsx.SetCellValue("Sheet1", fmt.Sprintf("G%d", row + j), detail["status"])
            xlsx.SetCellValue("Sheet1", fmt.Sprintf("H%d", row + j), detail["comment"])
            xlsx.SetCellStyle("Sheet1", fmt.Sprintf("B%d", row + j), fmt.Sprintf("H%d", row + j), style)
        }


        xlsx.MergeCell("sheet1", fmt.Sprintf("A%d", row), fmt.Sprintf("A%d", row + len(details) - 1))
        xlsx.MergeCell("sheet1", fmt.Sprintf("B%d", row), fmt.Sprintf("B%d", row + len(details) - 1))
        xlsx.MergeCell("sheet1", fmt.Sprintf("C%d", row), fmt.Sprintf("C%d", row + len(details) - 1))

        // xlsx.SetColWidth("Sheet1", "A", "C", 80)



        row = row + len(details)


    }












    // Set active sheet of the workbook.
    xlsx.SetActiveSheet(1)
    // Save xlsx file by the given path.
    res := xlsx.SaveAs("./Workbook.xlsx")
    if res != nil {
        fmt.Println(res)

    }

}





