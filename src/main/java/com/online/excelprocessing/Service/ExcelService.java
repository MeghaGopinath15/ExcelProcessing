package com.online.excelprocessing.Service;

import com.online.excelprocessing.Model.Product;
//import com.online.excelprocessing.ProductRepository;
import com.online.excelprocessing.Repository.ProductRepository;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@Service
public class ExcelService {

        @Autowired
        private ProductRepository productRepository;

        public void processExcelFile(MultipartFile file) throws IOException {
                // Use Apache POI to read Excel data
                Workbook workbook = new XSSFWorkbook(file.getInputStream());
                Sheet sheet = workbook.getSheetAt(0);

                // Iterate through the rows
                for (Row row : sheet) {
                        Product product = new Product();
                        product.setProductName(row.getCell(1).getStringCellValue());
                        product.setPrice(row.getCell(2).getNumericCellValue());
                        product.setCategory(row.getCell(3).getStringCellValue());
                        product.setStock((int) row.getCell(4).getNumericCellValue());

                        productRepository.save(product);
                }
                workbook.close();
        }

}