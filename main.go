package main

import (
	"database/sql"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize" //_ "github.com/jackc/pgx"
	"github.com/jmoiron/sqlx"
	_ "github.com/lib/pq"
)

type excelBook struct {
	prefics string
}

type goodsGroup struct {
	Id        int
	GroupName sql.NullString
}

var currentBook excelBook = excelBook{prefics: "price-"}
var goodsGroups = []goodsGroup{}

//var xlsx *excelize.File

func main() {

	//fmt.Println("Начинаем формировать Прайс-лист")
	connStr := "user=postgres password=leto2015 dbname=postgres sslmode=disable"
	db, err := sqlx.Connect("postgres", connStr)
	if err != nil {
		panic(err)
	}
	defer db.Close()

	err = db.Select(&goodsGroups, "select a.id, a.gg_name as GroupName from public.goods_group as a where a.id > 0")
	if err != nil {
		fmt.Println(err)
	}

	createCityPriceList("moscow")

	fmt.Println("Закончили формировать Прайс-лист")
}

func createCityPriceList(cityName string) {
	fmt.Println("Создали Прайс-лист для ", cityName)

	xlsx := excelize.NewFile()

	initBook(xlsx)
	addTableOfContent()
	addCategoryList(xlsx)
	err := xlsx.SaveAs("./" + currentBook.prefics + cityName + ".xlsx")
	if err != nil {
		fmt.Println(err)
	}
}

func initBook(xlsx *excelize.File) {
	fmt.Println("Задаем настройки книги")
	err := xlsx.SetDocProps(&excelize.DocProperties{
		Category:       "category",
		ContentStatus:  "Draft",
		Created:        "2019-06-04T22:00:10Z",
		Creator:        "Go Excelize",
		Description:    "This file created by Go Excelize",
		Identifier:     "xlsx",
		Keywords:       "Spreadsheet",
		LastModifiedBy: "Go Author",
		Modified:       "2019-06-04T22:00:10Z",
		Revision:       "0",
		Subject:        "Test Subject",
		Title:          "Test Title",
		Language:       "en-US",
		Version:        "1.0.0",
	})

	if err != nil {
		fmt.Println(err)
	}
	/*
		PHPExcel_Settings::setLocale('ru_ru');

		$this->excel->getDefaultStyle()->applyFromArray([
			'font' => [
				'size' => 10,
			],
			'alignment' => [
				'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER,
			],
		]);
		// -- -- -- --

		// -- Записываем информацию в заголовок Excel-файла
		$this->excel->getProperties()
			->setCreator($this->site->getFrontendDomain())
			->setLastModifiedBy($this->site->getFrontendDomain())
			->setTitle(' - прайс лист на ' . date('d.m.Y'))
			->setSubject(' ')
			->setDescription('')
			->setKeywords('')
			->setCategory('')
		;
	*/
}

func addTableOfContent() {
	fmt.Println("Добавляем оглавление")
	//$this->tableOfContent = new ExcelSheet_TableOfContent($this, $this->excel->setActiveSheetIndex(0));
}

func addCategoryList(xlsx *excelize.File) {
	fmt.Println("Добавляем листы категорий")
	for _, group := range goodsGroups {
		if len(group.GroupName.String) > 0 {
			xlsx.NewSheet(group.GroupName.String)
		}
		fmt.Println("Добавляем листы категорий " + group.GroupName.String)
	}
}
