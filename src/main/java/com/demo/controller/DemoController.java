package com.demo.controller;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RequestPart;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;


@RestController
public class DemoController {

	@GetMapping("/")
	public String index() {
		return "Greetings fr!";
	}
	
	public static boolean isRowEmpty(Row row) {
		boolean isEmpty = true;
		DataFormatter dataFormatter = new DataFormatter();
		if (row != null) {
			for (Cell cell : row) {
				if (dataFormatter.formatCellValue(cell).trim().length() > 0) {
					isEmpty = false;
					break;
				}
			}
		}
		return isEmpty;
	}
	
	@PostMapping("/upd")
	public void up ( @RequestPart("file") MultipartFile file ) throws IOException {
	    try{
            Workbook workbook = new XSSFWorkbook(file.getInputStream());
            Sheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rows = sheet.iterator();
            int rowNumber = 0;
            while ( rows.hasNext()) {
                Row currentRow = rows.next();
                if (isRowEmpty(currentRow)) {
                    continue;
                }
                if (rowNumber == 0) {
                    rowNumber++;
                    continue;
                }
                Iterator<Cell> cellsInRow = currentRow.iterator();
                String col1 = null;
                String col2 = null;
                int cellIdx = 0;
                DataFormatter formatter = new DataFormatter();
                while (cellsInRow.hasNext() || cellIdx < 2) {
                    if (cellsInRow.hasNext()) {
                        Cell currentCell = cellsInRow.next();
                    }                   
                    Cell cell = currentRow.getCell(cellIdx);
                    switch (cellIdx) {
                        case 0:if(cell!= null) {
                            if(cell.getCellType() != CellType.BLANK) {
                                col1= formatter.formatCellValue(cell).trim();
                            }
                        }break;
                        case 1:if(cell!= null) {
                            if(cell.getCellType() != CellType.BLANK) {
                                col2= formatter.formatCellValue(cell).trim();
                            }
                        }break;
                        default:break;
                    }
                    cellIdx++;
                } 
                System.out.println("Column 1 : "+col1+"  =====   column 2 : "+col2);
            }

        }
		catch ( Exception e ) {
            e.printStackTrace();
            System.out.println(e.getMessage() );
        }
 	}
}
