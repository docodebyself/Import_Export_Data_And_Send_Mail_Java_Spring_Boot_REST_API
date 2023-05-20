package com.example.exportdatatoexcel.utils;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import javax.swing.undo.CannotUndoException;
import java.util.ArrayList;
import java.util.List;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class ImportConfig {

    private int sheetIndex;

    private int headerIndex;

    private int startRow;

    private Class dataClazz;

    private List<CellConfig> cellImportConfigs;


    public static final ImportConfig customerImport;

    static{
        customerImport = new ImportConfig();
        customerImport.setSheetIndex(0);
        customerImport.setHeaderIndex(0);
        customerImport.setStartRow(1);
        customerImport.setDataClazz(CustomerDTO.class);
        List<CellConfig> customerImportCellConfigs = new ArrayList<>();

        customerImportCellConfigs.add(new CellConfig(0, "customerNumber"));
        customerImportCellConfigs.add(new CellConfig(1, "customerName"));
        customerImportCellConfigs.add(new CellConfig(2, "contactLastName"));
        customerImportCellConfigs.add(new CellConfig(3, "contactFirstName"));
        customerImportCellConfigs.add(new CellConfig(4, "phone"));
        customerImportCellConfigs.add(new CellConfig(5, "addressLine1"));
        customerImportCellConfigs.add(new CellConfig(6, "addressLine2"));
        customerImportCellConfigs.add(new CellConfig(7, "city"));
        customerImportCellConfigs.add(new CellConfig(8, "state"));
        customerImportCellConfigs.add(new CellConfig(9, "postalCode"));
        customerImportCellConfigs.add(new CellConfig(10, "country"));
        customerImportCellConfigs.add(new CellConfig(11, "salesRepEmployeeNumber"));
        customerImportCellConfigs.add(new CellConfig(12, "creditLimit"));

        customerImport.setCellImportConfigs(customerImportCellConfigs);
    }

}
