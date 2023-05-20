package com.example.exportdatatoexcel.service.impl;

import com.example.exportdatatoexcel.entity.Customer;
import com.example.exportdatatoexcel.repository.CustomerRepository;
import com.example.exportdatatoexcel.service.CustomerService;
import com.example.exportdatatoexcel.service.MailService;
import com.example.exportdatatoexcel.utils.*;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.http.HttpStatus;
import org.springframework.stereotype.Service;
import org.springframework.util.CollectionUtils;
import org.springframework.web.multipart.MultipartFile;

import java.util.List;

@Service
@RequiredArgsConstructor
public class CustomerServiceImpl implements CustomerService {

    private final CustomerRepository customerRepository;

    private final MailService mailService;

    @Override
    public BaseResponse importCustomerData(MultipartFile importFile) {


        BaseResponse baseResponse = new BaseResponse();

        Workbook workbook = FileFactory.getWorkbookStream(importFile);

        List<CustomerDTO> customerDTOList = ExcelUtils.getImportData(workbook, ImportConfig.customerImport);

        if(!CollectionUtils.isEmpty(customerDTOList)){
            saveData(customerDTOList);
            baseResponse.setCode(String.valueOf(HttpStatus.OK));
            baseResponse.setMessage("Import successfully");
        }else{
            baseResponse.setCode(String.valueOf(HttpStatus.BAD_REQUEST));
            baseResponse.setMessage("Import failed");
        }

        return baseResponse;
    }

    @Override
    public BaseResponse createCustomer(CustomerDTO customerDTO) {

        BaseResponse baseResponse = new BaseResponse();
        baseResponse.setCode("200");
        baseResponse.setMessage("Create successfully");

        //saveData(List.of(customerDTO));

        //send mail
        mailService.sendMailCreateCustomer(customerDTO);
        return baseResponse;
    }

    private void saveData(List<CustomerDTO> customerDTOList){
        for(CustomerDTO customerDTO : customerDTOList){
            Customer customer = new Customer();
            customer.setCustomerNumber(customerDTO.getCustomerNumber());
            customer.setCustomerName(customerDTO.getCustomerName());
            customer.setContactFirstName(customerDTO.getContactFirstName());
            customer.setContactLastName(customerDTO.getContactLastName());
            customer.setPhone(customerDTO.getPhone());
            customer.setAddressLine1(customerDTO.getAddressLine1());
            customer.setAddressLine2(customerDTO.getAddressLine2());
            customer.setCity(customerDTO.getCity());
            customer.setState(customerDTO.getState());
            customer.setPostalCode(customerDTO.getPostalCode());
            customer.setCountry(customerDTO.getCountry());
            customer.setSalesRepEmployeeNumber(customerDTO.getSalesRepEmployeeNumber());
            customer.setCreditLimit(customerDTO.getCreditLimit());
            customerRepository.save(customer);
        }
    }
}
