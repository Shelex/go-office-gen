package main

import (
	"fmt"
	"log"
	"strconv"

	"github.com/nguyenthenguyen/docx"
	"github.com/xuri/excelize/v2"
)

type Person struct {
	Name    string
	Age     int
	Address string
}

func prepareInvitation(templateFile *docx.ReplaceDocx, person Person) error {
	template := templateFile.Editable()
	template.Replace(`.NAME_TEMPLATE`, person.Name, -1)
	template.Replace(".AGE_TEMPLATE", strconv.Itoa(person.Age), -1)
	template.Replace(".ADDRESS_TEMPLATE", person.Address, -1)
	return template.WriteToFile(fmt.Sprintf("output/%s_invite.docx", person.Name))
}

func readPersons(input *excelize.File) ([]Person, error) {
	rows, err := input.GetRows("Sheet1")
	if err != nil {
		return nil, fmt.Errorf("failed to get data file row: %s", err)
	}

	persons := make([]Person, len(rows))

	for idx, row := range rows {
		age, err := strconv.Atoi(row[1])
		if err != nil {
			log.Printf("failed to convert age cell to number: %s", err)
		}

		persons[idx] = Person{
			Name:    row[0],
			Age:     age,
			Address: row[2],
		}
	}
	return persons, nil
}

func main() {
	input, err := excelize.OpenFile("input.xlsx")
	if err != nil {
		log.Printf("failed to read data file: %s", err)
		return
	}

	defer func() {
		if err := input.Close(); err != nil {
			log.Printf("failed to close data file: %s", err)
		}
	}()

	templateFile, err := docx.ReadDocxFile("./Template.docx")
	if err != nil {
		log.Println("failed to read template")
	}

	defer func() {
		if err := templateFile.Close(); err != nil {
			log.Printf("failed to close template file: %s", err)
		}
	}()

	persons, err := readPersons(input)
	if err != nil {
		log.Printf("failed to read persons: %s", err)
		return
	}

	for _, person := range persons {
		if err := prepareInvitation(templateFile, person); err != nil {
			log.Printf("failed to prepare invitation for %s:%s", person.Name, err)
		} else {
			log.Printf("saved invitation for %s", person.Name)
		}
	}
}
