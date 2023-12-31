package main

import (
	"flag"
	"fmt"
	"github.com/xuri/excelize/v2"
	"image"
	"image/color"
	_ "image/jpeg"
	_ "image/png"
	"log"
	"os"
	"strconv"
	"sync"
)

var colorIndex int

func main() {
	// 图片路径
	var filePath string
	// 每格多少像素
	var pix int
	// 是否输出颜色序号
	var isPrint bool
	// 生成excel格子高度
	var cellHeight float64

	flag.StringVar(&filePath, "path", "", "file path")
	flag.BoolVar(&isPrint, "print", false, "print color number")
	flag.IntVar(&pix, "pix", 16, "pix")
	flag.Float64Var(&cellHeight, "cellHeight", 20, "excel cell height")
	flag.Parse()
	f, err := os.Open(filePath)
	if err != nil {
		log.Printf("open file error:%v", err)
		panic(err)
	}
	defer f.Close()
	img, _, err := image.Decode(f)
	if err != nil {
		log.Printf("decode image error:%v", err)
		panic(err)
	}
	// 图片宽高
	length := img.Bounds().Max.X
	height := img.Bounds().Max.Y
	// 构造excel列坐标，最高支持10000列
	xMap := initX()
	// 创建excel文档
	ef := excelize.NewFile()
	defer func() {
		if err := ef.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	// Create a new sheet.
	index, err := ef.NewSheet("Sheet1")
	if err != nil {
		log.Printf("create new sheet error:%v", err)
		panic(err)
	}
	// Set active sheet of the workbook.
	ef.SetActiveSheet(index)
	colorMap := make(map[string]int)
	colorIndex = 0
	lengthNum := length / pix
	if length%pix != 0 {
		lengthNum++
	}
	heightNum := height / pix
	if height%pix != 0 {
		height++
	}
	var wg sync.WaitGroup
	var mutex sync.Mutex
	// 遍历横格子
	for i := 0; i < lengthNum; i++ {
		// 遍历竖格子
		for j := 0; j < heightNum; j++ {
			indexI := i
			indexJ := j
			wg.Add(1)
			go process(&wg, &mutex, indexI, indexJ, pix, img, colorMap, ef, xMap, cellHeight)
		}
	}
	wg.Wait()
	// 设置列宽
	_ = ef.SetColWidth("Sheet1", xMap[1], xMap[lengthNum], cellHeight*0.3528/2.2733)
	// Save spreadsheet by the given path.
	if err := ef.SaveAs("./result.xlsx"); err != nil {
		log.Printf("save excel file error:%v", err)
		panic(err)
	}
	// 输出颜色序号列表
	if isPrint {
		for k, v := range colorMap {
			fmt.Println(k + ":" + strconv.Itoa(v))
		}
	}
}

// 生成excel列坐标
func initX() map[int]string {
	rs := make(map[int]string)
	for i := 1; i <= 10000; i++ {
		x := columnName(i)
		rs[i] = x
	}
	return rs
}

func columnName(n int) string {
	colName := ""
	for n > 0 {
		n--
		colName = string(rune('A'+(n%26))) + colName
		n /= 26
	}
	return colName
}

func process(wg *sync.WaitGroup, mutex *sync.Mutex, i int, j int, pix int, img image.Image, colorMap map[string]int, ef *excelize.File, xMap map[int]string, cellHeight float64) {
	defer wg.Done()
	num := 0
	// 计算该格子像素点
	x := i*pix + pix/2
	y := j*pix + pix/2
	r, g, b, a := img.At(x, y).RGBA()
	rgba := color.RGBA{R: uint8(r >> 8), G: uint8(g >> 8), B: uint8(b >> 8), A: uint8(a >> 8)}
	colorStr := rgbToHex(rgba.R, rgba.G, rgba.B)
	// 加锁
	mutex.Lock()
	defer mutex.Unlock()
	v, ok := colorMap[colorStr]
	if !ok {
		// 该颜色第一次出现
		colorIndex++
		colorMap[colorStr] = colorIndex
		num = colorIndex
	} else {
		num = v
	}
	// 写入excel
	style, _ := ef.NewStyle(&excelize.Style{
		Border: []excelize.Border{
			{Type: "top", Color: "#000000", Style: 1},
			{Type: "bottom", Color: "#000000", Style: 1},
			{Type: "left", Color: "#000000", Style: 1},
			{Type: "right", Color: "#000000", Style: 1},
		},
		Fill: excelize.Fill{Type: "pattern", Color: []string{colorStr}, Pattern: 1},
	})
	_ = ef.SetCellValue("Sheet1", xMap[i+1]+strconv.Itoa(j+1), num)
	_ = ef.SetCellStyle("Sheet1", xMap[i+1]+strconv.Itoa(j+1), xMap[i+1]+strconv.Itoa(j+1), style)
	// 设置行高
	_ = ef.SetRowHeight("Sheet1", j+1, cellHeight)
}

// rgb转16进制
func rgbToHex(r, g, b uint8) string {
	return fmt.Sprintf("#%02X%02X%02X", r, g, b)
}
