package com.example.exportdatatoexcel.utils;

import com.example.exportdatatoexcel.entity.Customer;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.ArrayList;
import java.util.List;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class ExportConfig {

    private int sheetIndex;

    private int startRow;

    private Class dataClazz;

    private List<CellConfig> cellExportConfigList;


    public static final ExportConfig customerExport;
    static{
        customerExport = new ExportConfig();
        customerExport.setSheetIndex(0);
        customerExport.setStartRow(1);
        customerExport.setDataClazz(Customer.class);
        List<CellConfig> customerCellConfig = new ArrayList<>();
        customerCellConfig.add(new CellConfig(0, "customerNumber"));
        customerCellConfig.add(new CellConfig(1, "customerName"));
        customerCellConfig.add(new CellConfig(2, "contactLastName"));
        customerCellConfig.add(new CellConfig(3, "contactFirstName"));
        customerCellConfig.add(new CellConfig(4, "phone"));
        customerCellConfig.add(new CellConfig(5, "addressLine1"));
        customerCellConfig.add(new CellConfig(6, "addressLine2"));
        customerCellConfig.add(new CellConfig(7, "city"));
        customerCellConfig.add(new CellConfig(8, "state"));
        customerCellConfig.add(new CellConfig(9, "postalCode"));
        customerCellConfig.add(new CellConfig(10, "country"));
        customerCellConfig.add(new CellConfig(11, "salesRepEmployeeNumber"));
        customerCellConfig.add(new CellConfig(12, "creditLimit"));

        customerExport.setCellExportConfigList(customerCellConfig);
    }

}
