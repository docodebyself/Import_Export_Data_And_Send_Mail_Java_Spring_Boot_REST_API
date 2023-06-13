package com.example.exportdatatoexcel.entity.dto;

import lombok.Data;

@Data
public class CustomerDTO {
    private Integer customerNumber;

    private String customerName;

    private String contactLastName;

    private String contactFirstName;

    private String phone;

    private String addressLine1;

    private String addressLine2;

    private String city;

    private String state;

    private String postalCode;

    private String country;

    private Integer salesRepEmployeeNumber;

    private Long creditLimit;
}
