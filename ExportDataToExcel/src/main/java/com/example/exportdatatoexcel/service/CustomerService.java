package com.example.exportdatatoexcel.service;

import com.example.exportdatatoexcel.utils.BaseResponse;
import com.example.exportdatatoexcel.utils.CustomerDTO;
import org.springframework.web.multipart.MultipartFile;

public interface CustomerService {
    BaseResponse importCustomerData(MultipartFile importFile);

    BaseResponse createCustomer(CustomerDTO customerDTO);
}
