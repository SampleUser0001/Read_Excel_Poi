package ittimfn.sample.poi.read.controller;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import ittimfn.sample.poi.read.model.ExcelModel;
import lombok.Data;

@Data
public class ReadExcelController {

    private FileInputStream stream;

    private Workbook workbook;
    private Sheet sheet;

    public ReadExcelController(String filepath) throws EncryptedDocumentException, FileNotFoundException, IOException {
        this.stream = new FileInputStream(filepath);
        this.workbook = WorkbookFactory.create(this.stream);
        this.sheet = workbook.getSheet("Sheet1");
    }

    public ExcelModel getModel() {
        ExcelModel model = new ExcelModel();

        Row row = sheet.getRow(0);
        Cell cell = row.getCell(0);

        model.setCellValue(cell.getStringCellValue());

        return model;
    }

    public void close() throws IOException {
        this.workbook.close();
        this.stream.close();
    }


}
