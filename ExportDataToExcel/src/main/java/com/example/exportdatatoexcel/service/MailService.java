package com.example.exportdatatoexcel.service;

import com.example.exportdatatoexcel.utils.CustomerDTO;

public interface MailService {
    void sendMailTest();

    void sendMailCreateCustomer(CustomerDTO dto);
}
