package com.qzx.excel.utils;

import com.alibaba.excel.write.handler.SheetWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;

import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

public class SelfWriteHandle implements SheetWriteHandler {

    private Map<Integer, String[]> mapDropDown;

    public SelfWriteHandle(Map<Integer, String[]> mapDropDown) {
        this.mapDropDown = mapDropDown;
    }

    @Override
    public void beforeSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {
    }

    /**
     * 增加单元格筛选限制条件（导出的excel增加下拉菜单）
     * */
    @Override
    public void afterSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {
        final Workbook workbook = writeWorkbookHolder.getWorkbook();
        Sheet sheet = writeSheetHolder.getSheet();
        // 开始设置下拉框
        DataValidationHelper helper = sheet.getDataValidationHelper();
        for (Map.Entry<Integer, String[]> entry : mapDropDown.entrySet()) {
            // 定义sheet的名称
            String sheetName = "PrivateSheetHidden" + entry.getKey();
            // 1.创建一个隐藏的sheet 名称为 privateSheet
            Sheet privateSheet = workbook.createSheet(sheetName);
            // 设置隐藏
            workbook.setSheetHidden(workbook.getSheetIndex(sheetName), true);
            // 2.循环赋值（为了防止下拉框的行数与隐藏域的行数相对应，将隐藏域加到结束行之后）
            // 设置下拉框数据
            List<String> values = Arrays.stream(entry.getValue()).collect(Collectors.toList());
            for (int i = 0, length = values.size(); i < length; i++) {
                // i:表示你开始的行数 0表示你开始的列数
                privateSheet.createRow(i).createCell(0).setCellValue(values.get(i));
            }
            Name category1Name = workbook.createName();
            category1Name.setNameName(sheetName);
            // 4 $A$1:$A$N代表 以A列1行开始获取N行下拉数据
            category1Name.setRefersToFormula(sheetName + "!$A$1:$A$" + (values.size()));
            // 5 将刚才设置的sheet引用到你的下拉列表中 //起始行、终止行、起始列、终止列
            CellRangeAddressList addressList = new CellRangeAddressList(1, 65535, entry.getKey(), entry.getKey());
            DataValidationConstraint constraint = helper.createFormulaListConstraint(sheetName);
            DataValidation dataValidation = helper.createValidation(constraint, addressList);
            sheet.addValidationData(dataValidation);
        }
    }
}
