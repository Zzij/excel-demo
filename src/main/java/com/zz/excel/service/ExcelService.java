package com.zz.excel.service;

import com.zz.excel.entity.ExcelEntity;
import com.zz.excel.util.DateFormatUtil;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import javax.swing.text.html.parser.Entity;
import java.io.*;
import java.util.ArrayList;
import java.util.List;

@Service
public class ExcelService {

    public List<ExcelEntity> readExcel(MultipartFile file) {
        List<ExcelEntity> list = new ArrayList<>();
        if (file == null) {
            throw new RuntimeException("文件不为空");
        }
        String filename = file.getOriginalFilename();
        if (!filename.endsWith("xlsx")) {
            throw new RuntimeException("文件不为excel");
        }
        Workbook workbook = null;
        try {
            InputStream inputStream = file.getInputStream();
            workbook = new XSSFWorkbook(inputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }

        if (workbook == null) {
            throw new RuntimeException("为空");
        }
        int sheets = workbook.getNumberOfSheets();
        for (int i = 0; i < sheets; i++) {
            Sheet sheet = workbook.getSheetAt(i);
            System.out.println(sheet.getLastRowNum());
            for (int j = 0; j < sheet.getLastRowNum() + 1; j++) {
                Row row = sheet.getRow(j);
                ExcelEntity entity = new ExcelEntity();
                if (row.getCell(0) != null) {
                    Cell cell = row.getCell(0);
                    CellType cellTypeEnum = cell.getCellTypeEnum();

                    System.out.println(cellTypeEnum);
                    cell.setCellType(CellType.STRING);
                    entity.setName(cell.getStringCellValue());
                }
                if (row.getCell(1) != null) {
                    Cell cell = row.getCell(1);
                    cell.setCellType(CellType.STRING);
                    entity.setAge(Integer.parseInt(cell.getStringCellValue()));
                }
                if (row.getCell(2) != null) {
                    Cell cell = row.getCell(2);
                    cell.setCellType(CellType.STRING);
                    entity.setEmail(cell.getStringCellValue());
                }
                if (row.getCell(3) != null) {
                    Cell cell = row.getCell(3);
                    cell.setCellType(CellType.STRING);
                    entity.setPhone(cell.getStringCellValue());
                }
                list.add(entity);
            }
        }

        return list;
    }

    public void writeExcel(ExcelEntity entity) {
        OutputStream os = null;
        File file = new File("D:/test.xls");
        try {
            if (file.exists()) {
                return;
            }
            os = new FileOutputStream(file);
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheet = workbook.createSheet("第一个单元");
            for (int i = 0; i < 2; i++) {
                HSSFRow row = sheet.createRow(i);
                row.createCell(0).setCellValue(entity.getName());
                row.createCell(1).setCellValue(entity.getAge());
                row.createCell(2).setCellValue(entity.getEmail());
                row.createCell(3).setCellValue(entity.getPhone());
            }

            workbook.write(os);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (os != null) {
                try {
                    os.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }


    public HSSFWorkbook downloadExcel(ExcelEntity entity) {

        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("第一个单元");
        for (int i = 0; i < 2; i++) {
            System.out.println(111);
            HSSFRow row = sheet.createRow(i);
            row.createCell(0).setCellValue(entity.getName());
            row.createCell(1).setCellValue(entity.getAge());
            row.createCell(2).setCellValue(entity.getEmail());
            row.createCell(3).setCellValue(entity.getPhone());
        }
        return workbook;
    }

    /**
     *
     * 获取excel表格中的内容
     * @param cell  单元格
     * @return  字符串
     */
    public String getCellValue(Cell cell){
        if(cell == null){
            return null;
        }
        CellType cellTypeEnum = cell.getCellTypeEnum();
        String value = null;
        switch (cellTypeEnum){
            case STRING:
                value = cell.getStringCellValue();
                break;
            case NUMERIC:
                if(HSSFDateUtil.isCellDateFormatted(cell)){
                    value = DateFormatUtil.getStrDate(cell.getDateCellValue());
                }else{
                    //如果按数字来读的话11 会读成11.0 所以转为string
                    cell.setCellType(CellType.STRING);
                    value = cell.getStringCellValue();
                }
                break;
            case _NONE:
                value = null;
                break;
            case BOOLEAN:
                value = String.valueOf(cell.getBooleanCellValue());
                break;
            case BLANK:
                break;
            case ERROR:
                value = "错误信息，无法读取";
                break;
            case FORMULA:
                value = cell.getCellFormula();
                break;
            default:
                value = "无法识别内容";
        }
        return value;
    }
}
