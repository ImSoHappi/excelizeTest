package main

import (
	"fmt"
	"math/rand"

	"github.com/xuri/excelize/v2"
)

func main() {
	f := excelize.NewFile()
	f.NewSheet("Sheet2")

	region := []string{"IX de la Araucanía", "RM Metropolitana", "V de Valparaíso", "VI del Libertador Gral. Bernardo O'higgins", "VII del Maule"}
	ciudad := []string{"Renaico", "Temuco", "Paine", "Las Cabras", "Rancagua", "San Fernando", "Palmilla", "Linares", "Molina", "Teno"}
	catAbuelo := []string{"Consumo Doméstico", "Equipos para Operación", "Ferretería", "Insumos Especificos de Acuicultura y Pesca", "Materiales Para Operación", "Packaging"}
	catPadre := []string{"Artículos de Aseo", "Farmacia", "Neumáticas", "Repuestos y Accesorios de Equipos para Operación", "Electricidad", "Fijaciones", "Gasfitería", "Herramientas", "Pinturas y Herramientas", "Quincallería y Cerraduras"}
	categoria := []string{"Otros Artículos de Aseo", "Farmacia y Cuidado Personal", "Accesorios y Herramientas Neumáticas", "Neumáticos", "Accesorios Instalación Eléctricos y Conectores", "Tornillos, Clavos, Pernos y otros", "Teflón y Soldaduras", "Herramientas Manuales"}
	producto := []string{"escobilla acero c/mango", "Huaipe", "PARCHE 75X45", "PARCHE REPARACION FRIO", "CINTA AISLANTE ROJA", "REMACHE POP 4 X 16 MM", "REMACHE POP 4.8 X 16 MM", "DISCO DESBASTE FIERRO 4 1/2", "DISCO TRASLAPADO 4 ½ GRANO 4", "Brocha 2", "CANDADO ODIS # 330 30MM", "CEMENTO VULCANIZADOR", "SIKADUR 32", "SILICONA AUTOMOTRIZ", "CORDEL PLASTICO TIPO MIL", "REGULADOR GAS"}

	f.SetSheetRow("Sheet2", "A1", &[]string{"Region", "Ciudad", "catAbuelo", "catPadre", "Categoria", "Producto", "Promedio"})

	for i := 0; i < 20; i++ {
		f.SetCellValue("Sheet2", fmt.Sprintf("A%d", i+2), region[rand.Intn(len(region))])
		f.SetCellValue("Sheet2", fmt.Sprintf("B%d", i+2), ciudad[rand.Intn(len(ciudad))])
		f.SetCellValue("Sheet2", fmt.Sprintf("C%d", i+2), catAbuelo[rand.Intn(len(catAbuelo))])
		f.SetCellValue("Sheet2", fmt.Sprintf("D%d", i+2), catPadre[rand.Intn(len(catPadre))])
		f.SetCellValue("Sheet2", fmt.Sprintf("E%d", i+2), categoria[rand.Intn(len(categoria))])
		f.SetCellValue("Sheet2", fmt.Sprintf("F%d", i+2), producto[rand.Intn(len(producto))])
		f.SetCellValue("Sheet2", fmt.Sprintf("G%d", i+2), float64(rand.Intn((10-1)/0.05))*0.05+1)
	}

	if err := f.AddPivotTable(&excelize.PivotTableOption{
		DataRange:        "Sheet2!$A$1:$G$21",
		PivotTableRange:  "Sheet1!$B$2:$M$34",
		Columns:          []excelize.PivotTableField{{Data: "Region", Name: "Region"}, {Data: "Ciudad", Name: "Ciudad"}},
		Rows:             []excelize.PivotTableField{{Data: "catAbuelo", Name: "Categorias"}, {Data: "catPadre"}, {Data: "Categoria"}, {Data: "Producto"}},
		Data:             []excelize.PivotTableField{{Data: "Promedio", Name: "Promedio de Of/Prod Simple", Subtotal: "Average"}},
		CompactData:      true,
		ShowDrill:        true,
		ShowRowHeaders:   true,
		ShowColHeaders:   true,
		PageOverThenDown: true,
		ShowLastColumn:   true,
		ShowRowStripes:   true,

		PivotTableStyleName: "PivotStyleMedium9",
	}); err != nil {
		fmt.Println(err)
	}

	if err := f.SaveAs("test.xlsx"); err != nil {
		fmt.Println(err)
	}
}
