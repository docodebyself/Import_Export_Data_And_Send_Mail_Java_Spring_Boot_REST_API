package com.example.exportdatatoexcel.service;

import com.example.exportdatatoexcel.utils.BaseResponse;
import com.example.exportdatatoexcel.utils.CustomerDTO;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.util.List;

public interface CustomerService {
    BaseResponse importCustomerData(MultipartFile importFile);

    BaseResponse createCustomer(CustomerDTO customerDTO);

    List<File> getFilesFromServer();

    List<File> zipExcelFileFromDatabase() throws Exception;
}
