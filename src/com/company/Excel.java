package com.company;




import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

public class Excel {
    private List<String> headers = new ArrayList<>();
    private List<File> errorExcels = new ArrayList<>();
    private int colCount;
    private Outer outer = new Outer();
    public int collect(Searcher s){
        File module = s.getModule();
        List<File> excels = s.getExcels();
        try{
            //先读取表头
            FileInputStream in = new FileInputStream(module);

            XSSFWorkbook moudleReadBook = new XSSFWorkbook(in);
            XSSFSheet moudleReadSheet = moudleReadBook.getSheetAt(0);
            XSSFRow row = moudleReadSheet.getRow(0);
            //获取模板列数和行数
            this.colCount = row.getPhysicalNumberOfCells();
            int lastMoudleRow = moudleReadSheet.getLastRowNum();
            //先判断一下module中是否有表头（row>-1）
            if(lastMoudleRow > -1){
                for(Cell c : row){
                    String value = null;
                    switch (c.getCellTypeEnum()){
                        case STRING:    value = c.getStringCellValue();break;
                        default:
                            System.out.println("module表头格式错误！(仅支持字符串)");
                            System.exit(1);
                    }
                    headers.add(value);
                }
            }

            //从待合并的excel中提取数据并加到模板中
            int successCount = 0;
            int lastMRow = moudleReadSheet.getLastRowNum();
            for(File f : excels){
                FileInputStream fin = new FileInputStream(f);
                XSSFWorkbook excelReadBook = new XSSFWorkbook(fin);
                XSSFSheet excelReadSheet = excelReadBook.getSheetAt(0);
                XSSFRow excelHeadRow = excelReadSheet.getRow(0);
                if (isHeadCorrected(excelHeadRow)){
                    int lastExcelRow = excelReadSheet.getLastRowNum();
                    //循环处理行
                    for(int i = 1; i <= lastExcelRow; i++){
                       XSSFRow excelRow = excelReadSheet.getRow(i);
                       List<String> contants = getContants(excelRow);
                       //在模板中创建新的一行
                       XSSFRow moduleWriteRow = moudleReadSheet.createRow(moudleReadSheet.getLastRowNum() + 1);
                        //循环处理列
                        for(int j = 0; j < contants.size(); j++){
                            moduleWriteRow.createCell(j).setCellValue(contants.get(j));
                        }
                    }
                    //统计成功合并的文件数量
                    successCount ++;
                }else{
                    //如果表头不一样，则将该文件插入错误列表中
                    errorExcels.add(f);
                }
                fin.close();
            }

            in.close();
            //输出文件
            FileOutputStream out = new FileOutputStream(module);
            out.flush();
            moudleReadBook.write(out);
            out.close();

            return this.outer.outCollectResult(successCount, this.errorExcels);
        }catch (Exception e){
            e.printStackTrace();
        }
        return -1;
    }

    private boolean isHeadCorrected(XSSFRow headRow) throws Exception{
        int i = 0;
        //先判断一下是否未空白表
        if(headRow != null){
            for (Cell c : headRow){
                String head = c.getStringCellValue();
                if(! head.equals(headers.get(i))){
                    return false;
                }
                i ++;
            }
            return true;
        }
        return false;

    }

    private List<String> getContants(XSSFRow row){
        DataFormatter formatter = new DataFormatter();
        List<String> contants = new ArrayList<>();
        for(int i = 0; i < this.colCount; i ++){
            Cell c = row.getCell(i);
            String value;
            if(c == null){
                value = "";
            }else{
                value = formatter.formatCellValue(c);
            }
            contants.add(value);
        }
        return contants;
    }

    public List<File> getErrorExcels() {
        return errorExcels;
    }
}






