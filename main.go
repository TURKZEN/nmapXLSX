package main

import (
	"encoding/xml"
	"fmt"
	"os"
	"strings"
	"github.com/tealeg/xlsx"
)

type NmapRun struct {
	XMLName xml.Name `xml:"nmaprun"`
	Hosts   []Host   `xml:"host"`
}

type Host struct {
	Addresses []Address `xml:"address"`
	Ports     []Port    `xml:"ports>port"`
}

type Address struct {
	Addr     string `xml:"addr,attr"`
	AddrType string `xml:"addrtype,attr"`
}

type Port struct {
	PortID   int    `xml:"portid,attr"`
	Protocol string `xml:"protocol,attr"`
	State    State  `xml:"state"`
	Service  Service `xml:"service"`
}

type Service struct {
	Name    string `xml:"name,attr"`
	Product string `xml:"product,attr"`
}

type State struct {
	State string `xml:"state,attr"`
}

func main() {
	if len(os.Args) != 3 {
		fmt.Println("Usage: nmapXLSX <xml_Nmap_Output> <Output>")
		return
	}

	inputFile := os.Args[1]
	outputFile := os.Args[2]

	// If the output file name does not end with ".xlsx", add it
	if !strings.HasSuffix(outputFile, ".xlsx") {
		outputFile += ".xlsx"
	}

	xmlData, err := os.ReadFile(inputFile)
	if err != nil {
		fmt.Printf("File reading error: %v\n", err)
		return
	}

	var report NmapRun
	err = xml.Unmarshal(xmlData, &report)
	if err != nil {
		fmt.Printf("XML parsing error: %v\n", err)
		return
	}

	// Excel file creation
	xlsxFile := xlsx.NewFile()
	sheet, err := xlsxFile.AddSheet("Open Ports")
	if err != nil {
		fmt.Printf("Excel sheet creation error: %v\n", err)
		return
	}

	// Header row
	headerRow := sheet.AddRow()
	headerRow.AddCell().SetValue("IP Address")
	headerRow.AddCell().SetValue("Port")
	headerRow.AddCell().SetValue("Protocol")
	headerRow.AddCell().SetValue("State")
	headerRow.AddCell().SetValue("Service")
	headerRow.AddCell().SetValue("Version")

	// Data rows
	for _, host := range report.Hosts {
		for _, port := range host.Ports {
			dataRow := sheet.AddRow()
			dataRow.AddCell().SetValue(host.Addresses[0].Addr) // Assuming the first address is used
			dataRow.AddCell().SetValue(port.PortID)
			dataRow.AddCell().SetValue(port.Protocol)
			dataRow.AddCell().SetValue(port.State.State)
			dataRow.AddCell().SetValue(port.Service.Name)
			dataRow.AddCell().SetValue(port.Service.Product)
		}
	}

	// Save the Excel file
	err = xlsxFile.Save(outputFile)
	if err != nil {
		fmt.Printf("Error saving Excel file: %v\n", err)
		return
	}

	fmt.Printf("Excel file created: %s \n", outputFile)
}
