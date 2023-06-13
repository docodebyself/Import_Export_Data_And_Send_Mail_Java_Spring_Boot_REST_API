package com.example.exportdatatoexcel.utils;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;
import org.springframework.util.ResourceUtils;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;

@Component
public class FileFactory {
    public static final String PATH_TEMPLATE = "F:\\template\\";

    public static File createFile(String fileName, Workbook workbook) throws Exception {

        workbook = new XSSFWorkbook();

        OutputStream out = new FileOutputStream(PATH_TEMPLATE + fileName);

        workbook.write(out);

        return ResourceUtils.getFile(PATH_TEMPLATE + fileName);
    }


    public static Workbook getWorkbookStream(MultipartFile importFile){
        InputStream inputStream = null;
        try {
            inputStream = importFile.getInputStream();

            Workbook workbook = WorkbookFactory.create(inputStream);

            return workbook;
        }catch (Exception e){
            e.printStackTrace();
            return null;
        }
    }

}
