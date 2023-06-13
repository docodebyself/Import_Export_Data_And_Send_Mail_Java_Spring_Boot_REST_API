package com.example.exportdatatoexcel.controller;

import com.example.exportdatatoexcel.entity.Customer;
import com.example.exportdatatoexcel.repository.CustomerRepository;
import com.example.exportdatatoexcel.service.CustomerService;
import com.example.exportdatatoexcel.utils.BaseResponse;
import com.example.exportdatatoexcel.utils.CustomerDTO;
import com.example.exportdatatoexcel.utils.ExcelUtils;
import lombok.RequiredArgsConstructor;
import org.apache.poi.util.IOUtils;
import org.springframework.core.io.InputStreamResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.util.CollectionUtils;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

@RestController
@RequestMapping("/customer")
@RequiredArgsConstructor
public class CustomerController {

    private final CustomerRepository customerRepository;

    private final CustomerService customerService;

    @GetMapping("/export")
    public ResponseEntity<Resource> exportCustomer() throws Exception {
        List<Customer> customerList = customerRepository.findAll();

        if (!CollectionUtils.isEmpty(customerList)) {
            String fileName = "Customer Export" + ".xlsx";

            ByteArrayInputStream in = ExcelUtils.exportCustomer(customerList, fileName);

            InputStreamResource inputStreamResource = new InputStreamResource(in);

            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION,
                            "attachment; filename=" + URLEncoder.encode(fileName, StandardCharsets.UTF_8)
                    )
                    .contentType(MediaType.parseMediaType("application/vnd.ms-excel; charset=UTF-8"))
                    .body(inputStreamResource);
        } else {
            throw new Exception("No data");

        }
    }

    @PostMapping("/import")
    public ResponseEntity<BaseResponse> importCustomerData(@RequestParam("file") MultipartFile importFile) {
        BaseResponse baseResponse = customerService.importCustomerData(importFile);
        return new ResponseEntity<>(baseResponse, HttpStatus.OK);
    }


    @PostMapping("/create")
    public ResponseEntity<BaseResponse> createCustomer(@RequestBody CustomerDTO customerDTO){
        return new ResponseEntity<>(customerService.createCustomer(customerDTO), HttpStatus.OK);
    }

    @PostMapping(value = "/zip", produces = "application/zip")
    public ResponseEntity<StreamingResponseBody> zipFilesFromServer() throws Exception{
        List<File> files = customerService.getFilesFromServer();

        return ResponseEntity
                .ok()
                .header("Content-Disposition", "attachment; filename= Result_files.zip")
                .body(out ->{
                    var zipOutputStream = new ZipOutputStream(out);

                    //package files
                    Set<String> fileNameAdded = new HashSet<>();

                    for(File file : files){
                        //new zip entry and copying input stream with file, after that close stream
                        if(!fileNameAdded.contains(file.getName())){
                            zipOutputStream.putNextEntry(new ZipEntry(file.getName()));
                            FileInputStream fileInputStream = new FileInputStream(file);
                            IOUtils.copy(fileInputStream, zipOutputStream);

                            fileInputStream.close();
                            zipOutputStream.closeEntry();
                            fileNameAdded.add(file.getName());
                        }
                    }
                    zipOutputStream.close();
                });
    }

    @PostMapping(value = "/zip-v2", produces = "application/zip")
    public ResponseEntity<StreamingResponseBody> zipFilesStoreDataFromDatabase() throws Exception{
        List<File> files = customerService.zipExcelFileFromDatabase();

        return ResponseEntity
                .ok()
                .header("Content-Disposition", "attachment; filename= Excel_Customers.zip")
                .body(out ->{
                    var zipOutputStream = new ZipOutputStream(out);

                    //package files
                    Set<String> fileNameAdded = new HashSet<>();

                    for(File file : files){
                        //new zip entry and copying input stream with file, after that close stream
                        if(!fileNameAdded.contains(file.getName())){
                            zipOutputStream.putNextEntry(new ZipEntry(file.getName()));
                            FileInputStream fileInputStream = new FileInputStream(file);
                            IOUtils.copy(fileInputStream, zipOutputStream);

                            fileInputStream.close();
                            zipOutputStream.closeEntry();
                            fileNameAdded.add(file.getName());
                        }
                    }
                    zipOutputStream.close();
                });
    }



}
