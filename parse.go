package main

/*
	XSLX to MXT file parser
	This utilite uses xuri/excelize
	GNU GPL v2.0, 2022
	N.S. Minaev
*/

import (
	_ "text/template"
)

// MobaXTerm connection struct
type MobaXTerm struct {
	RecordID      string // Bookmark #
	RKName        string // Региональная компания
	AptName       string // Имя аптеки в нотации типа 'мскАпт1001'
	ServerAddress string // Непостоянная часть адреса сервера
	Username      string // пользователь системы, efarma по умолчанию
}

func main() {
	print("Let's go parse files...")

}
