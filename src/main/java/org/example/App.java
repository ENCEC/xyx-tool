package org.example;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.collection.ListUtil;
import cn.hutool.core.convert.Convert;
import cn.hutool.core.util.ObjectUtil;
import cn.hutool.core.util.StrUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.*;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

/**
 * Hello world!
 */
@Slf4j
public class App {
    public static void main(String[] args) throws IOException {
        String filePath = "E:\\source1.xlsx";
        String outFilePath = "E:\\target1.xlsx";
        int logisticColIndex = 9;
        int orderColIndex = 11;
        int[] costIndexArray = {15, 16, 17, 18, 19, 22};
        String[] wechatArray = {"微信", "代理"};
        //获取工作簿
        XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(filePath));
        //获取工作表
        XSSFSheet mainSheet = workbook.getSheet("Sheet1");
        //获取物流数据
        Map<String, BigDecimal> logisticsMap = getLogisticsMap(workbook, "Sheet5");
        //获取店铺数据
        Map<String, ShopInfo> shopMap = getShopMap(workbook, "Sheet4", 12, 13, 14);

        log.info("last row num: {}", mainSheet.getLastRowNum());
        log.info("last col num: {}", mainSheet.getRow(0).getLastCellNum());
        List<CellRangeAddress> mergeCellAddressList = getMergeCellList(mainSheet, logisticColIndex);
        //物流列合并单元格所占行下标集合
        List<Integer> mergeRowIndex = new ArrayList<>();
        mergeCellAddressList.forEach(e -> {
            IntStream.range(e.getFirstRow(), e.getLastRow() + 1).forEach(t -> mergeRowIndex.add(t));
        });

        List<Integer> multiValueOneRowIndex = getMultiValueOneCellList(mainSheet, logisticColIndex, mergeRowIndex);

        multiValueOneRowIndex.forEach(e -> log.info("multiValueOneRowIndex------:{}", e));
        mergeRowIndex.forEach(e -> log.info("mergeRowIndex------:{}", e));

        Row firstRow = mainSheet.getRow(0);
        short lastCellNum = firstRow.getLastCellNum();
        //创建表头
        String[] addCols = {"店铺", "成本", "物流", "提成", "推广", "金额", "扇数"};
        for (int i = 0; i < addCols.length; i++) {
            Cell cell = firstRow.createCell(lastCellNum + i);
            cell.setCellValue(addCols[i]);
        }
        dealMultiValueOneRowLogisticData(mainSheet, multiValueOneRowIndex, logisticsMap, logisticColIndex, lastCellNum + 2);
        dealMergeRowLogisticData(mainSheet, mergeCellAddressList, logisticsMap, logisticColIndex, lastCellNum + 2);
        List<Integer> allRowIndexList = IntStream.range(1, mainSheet.getLastRowNum() + 1).boxed().collect(Collectors.toList());
        Collection<Integer> subtractList1 = CollUtil.subtract(allRowIndexList, multiValueOneRowIndex);
        Collection<Integer> normalCollection = CollUtil.subtract(subtractList1, mergeRowIndex);
        dealNormalRowLogisticData(mainSheet, normalCollection, logisticsMap, logisticColIndex, lastCellNum + 2);
        //填充成本列数据
        fillCostData(mainSheet, allRowIndexList, costIndexArray, lastCellNum + 1);

        List<CellRangeAddress> shopMergeCellAddressList = getMergeCellList(mainSheet, orderColIndex);
        List<Integer> shopMergeRowIndex = new ArrayList<>();
        shopMergeCellAddressList.forEach(e -> {
            IntStream.range(e.getFirstRow(), e.getLastRow() + 1).forEach(t -> shopMergeRowIndex.add(t));
        });
        //填充订单店铺和金额列
        Collection<Integer> normalShopIndexData = CollUtil.subtract(allRowIndexList, shopMergeRowIndex);
        fillMergeOrderData(mainSheet, shopMergeCellAddressList, shopMap, orderColIndex, lastCellNum, lastCellNum + 5, wechatArray);
        fillNormalOrderData(mainSheet, normalShopIndexData, shopMap, orderColIndex, lastCellNum, lastCellNum + 5, wechatArray);

        //TODO 待修改
//        String cellValue1 = mainSheet.getRow(2).getCell(logisticColIndex).getStringCellValue();
//        String cellValue2 = mainSheet.getRow(32).getCell(logisticColIndex).getStringCellValue();
//        String[] array = cellValue1.split("\n");
//        String[] array2 = cellValue2.split("\n");
//        BigDecimal totalValue1 = BigDecimal.ZERO;
//        for (int i = 0; i < array.length; i++) {
//            //查找对应列的值
//            XSSFCell logisticsCell = mainSheet.getRow(2 + i).createCell(lastCellNum + 2);
//            logisticsCell.setCellValue(0L);
//            totalValue1 = totalValue1.add(Convert.toBigDecimal(logisticsMap.getOrDefault(array[i], 0L)));
//        }
//        XSSFCell logisticsCell1 = mainSheet.getRow(2).createCell(lastCellNum + 2);
//        logisticsCell1.setCellValue(Convert.toLong(totalValue1));
//
//        BigDecimal totalValue = BigDecimal.ZERO;
//        for (int i = 0; i < array2.length; i++) {
//            totalValue = totalValue.add(Convert.toBigDecimal(logisticsMap.getOrDefault(array2[i], 0L)));
//        }
//        //查找对应列的值
//        XSSFCell logisticsCell = mainSheet.getRow(32).createCell(lastCellNum + 2);
//        logisticsCell.setCellValue(Convert.toLong(totalValue));

//        log.info("合并单元格开始行：{}",CellUtil.getCellRangeAddress(mainSheet.getRow(2).getCell(logisticColIndex)).getFirstRow());
//        log.info("合并单元格结束行：{}",CellUtil.getCellRangeAddress(mainSheet.getRow(2).getCell(logisticColIndex)).getLastRow());

        try (FileOutputStream out = new FileOutputStream(outFilePath)) {
            workbook.write(out);
        }

    }

    /**
     * @return void
     * @Author chenec
     * @Date 2024/2/17 17:12
     * @Description 处理合并的订单列
     * @Param [mainSheet, normalShopIndexData, shopMap, shopColIndex, lastCellNum]
     **/
    private static void fillNormalOrderData(XSSFSheet mainSheet, Collection<Integer> normalShopIndexData, Map<String, ShopInfo> shopMap,
                                            int orderColIndex, int shopNameCellIndex, int fcyIndex, String[] wechatArray) {
        log.info("=======fillNormalOrderData==========");
        for (Integer rowIndex : normalShopIndexData) {
            XSSFCell cell = mainSheet.getRow(rowIndex).getCell(orderColIndex);
            if (null == cell) {
                continue;
            }
            String cellValue = StrUtil.trim(cell.getStringCellValue());
            XSSFCell shopNameCell = mainSheet.getRow(rowIndex).createCell(shopNameCellIndex);
            XSSFCell fcyCell = mainSheet.getRow(rowIndex).createCell(fcyIndex);
            ShopInfo shopInfo = shopMap.get(cellValue);
            String shopName = ObjectUtil.isNull(shopInfo) ? "" : shopInfo.getShopName();
            BigDecimal price = ObjectUtil.isNull(shopInfo) ? BigDecimal.ZERO : shopInfo.getPrice();
            shopNameCell.setCellValue(shopName);
            fcyCell.setCellValue(price.doubleValue());
            if (ObjectUtil.isNotNull(shopInfo) && shopInfo.getPrice().compareTo(BigDecimal.ZERO) < 0) {
                fcyCell.setCellValue("N/A");
            }
            if (Arrays.stream(wechatArray).anyMatch(e -> e.equals(cellValue))) {

            }
        }
    }

    /**
     * @return void
     * @Author chenec
     * @Date 2024/2/17 17:12
     * @Description 处理合并的订单列
     * @Param [mainSheet, shopMergeCellAddressList, shopMap, shopColIndex, lastCellNum]
     **/
    private static void fillMergeOrderData(XSSFSheet mainSheet, List<CellRangeAddress> shopMergeCellAddressList, Map<String, ShopInfo> shopMap,
                                           int orderColIndex, int shopNameCellIndex, int fcyIndex, String[] wechatArray) {
        log.info("=======fillNormalOrderData==========");
        for (CellRangeAddress cellRangeAddress : shopMergeCellAddressList) {
            int firstRow = cellRangeAddress.getFirstRow();
            int lastRow = cellRangeAddress.getLastRow();
            String cellValue = mainSheet.getRow(firstRow).getCell(orderColIndex).getStringCellValue();
            for (int i = firstRow; i <= lastRow; i++) {
                XSSFCell shopNameCell = mainSheet.getRow(i).createCell(shopNameCellIndex);
                XSSFCell fcyCell = mainSheet.getRow(i).createCell(fcyIndex);
                ShopInfo shopInfo = shopMap.get(cellValue);
                String shopName = ObjectUtil.isNull(shopInfo) ? "" : shopInfo.getShopName();
                shopNameCell.setCellValue(shopName);
                fcyCell.setCellValue(BigDecimal.ZERO.doubleValue());
            }
            XSSFCell fcyCell = mainSheet.getRow(firstRow).createCell(fcyIndex);
            ShopInfo shopInfo = shopMap.get(cellValue);
            BigDecimal price = ObjectUtil.isNull(shopInfo) ? BigDecimal.ZERO : shopInfo.getPrice();
            fcyCell.setCellValue(price.doubleValue());
            if (shopInfo.getPrice().compareTo(BigDecimal.ZERO) < 0) {
                fcyCell.setCellValue("N/A");
            }
            //（处理微信和代理的情况）
        }
    }

    /**
     * @return void
     * @Author chenec
     * @Date 2024/2/17 15:13
     * @Description 填充成本列数据
     * @Param [mainSheet, allRowIndexList, costIndexArray, i]
     **/

    private static void fillCostData(XSSFSheet mainSheet, List<Integer> allRowIndexList, int[] costIndexArray, int targetIndex) {
        for (Integer rowIndex : allRowIndexList) {
            BigDecimal totalValue = BigDecimal.ZERO;
            for (int colIndex : costIndexArray) {
                XSSFCell cell = mainSheet.getRow(rowIndex).getCell(colIndex);
                if (null == cell) {
                    continue;
                }
                log.info("{}", rowIndex);
                totalValue = totalValue.add(Convert.toBigDecimal(cell.getNumericCellValue()));
            }
            XSSFCell costCell = mainSheet.getRow(rowIndex).createCell(targetIndex);
            costCell.setCellValue(totalValue.doubleValue());
        }
    }

    /**
     * @return void
     * @Author chenec
     * @Date 2024/2/17 11:21
     * @Description 处理普通单元格的情况，物流数据查找填充
     * @Param [mainSheet, normalCollection, logisticsMap, logisticColIndex, i]
     **/
    private static void dealNormalRowLogisticData(XSSFSheet mainSheet, Collection<Integer> normalCollection, Map<String, BigDecimal> logisticsMap, int logisticColIndex, int targetCellIndex) {
        log.info("=======dealNormalRowLogisticData==========");
        for (Integer rowIndex : normalCollection) {
            XSSFCell cell = mainSheet.getRow(rowIndex).getCell(logisticColIndex);
            if (null == cell) {
                continue;
            }
            String cellValue = StrUtil.trim(cell.getStringCellValue());
            XSSFCell logisticsCell = mainSheet.getRow(rowIndex).createCell(targetCellIndex);
            logisticsCell.setCellValue(Convert.toDouble(logisticsMap.getOrDefault(cellValue, BigDecimal.ZERO)));
        }
    }

    /**
     * @return void
     * @Author chenec
     * @Date 2024/2/17 10:34
     * @Description 处理合并单元格的情况，物流数据查找填充
     * @Param [mainSheet, mergeRowIndex, logisticsMap, logisticColIndex, i]
     **/
    private static void dealMergeRowLogisticData(XSSFSheet mainSheet, List<CellRangeAddress> mergeCellAddressList, Map<String, BigDecimal> logisticsMap, int logisticColIndex, int targetCellIndex) {
        log.info("=======dealMergeRowLogisticData==========");
        for (CellRangeAddress cellRangeAddress : mergeCellAddressList) {
            int firstRow = cellRangeAddress.getFirstRow();
            int lastRow = cellRangeAddress.getLastRow();
            String cellValue = mainSheet.getRow(firstRow).getCell(logisticColIndex).getStringCellValue();
            String[] array = cellValue.split("\n");
            BigDecimal totalValue = BigDecimal.ZERO;
            for (int i = 0; i < array.length; i++) {
                totalValue = totalValue.add(Convert.toBigDecimal(logisticsMap.getOrDefault(StrUtil.trim(array[i]), BigDecimal.ZERO)));
            }
            for (int i = firstRow; i <= lastRow; i++) {
                XSSFCell logisticsCell = mainSheet.getRow(i).createCell(targetCellIndex);
                logisticsCell.setCellValue(0L);
            }
            XSSFCell logisticsCell1 = mainSheet.getRow(firstRow).createCell(targetCellIndex);
            logisticsCell1.setCellValue(Convert.toDouble(totalValue));
        }
    }

    /**
     * @return void
     * @Author chenec
     * @Date 2024/2/17 10:29
     * @Description 处理一个单元格内多值的情况，物流数据查找填充
     * @Param [mainSheet, multiValueOneRowIndex, logisticsMap, logisticColIndex, targetCellIndex]
     **/
    private static void dealMultiValueOneRowLogisticData(XSSFSheet mainSheet, List<Integer> multiValueOneRowIndex, Map<String, BigDecimal> logisticsMap, int logisticColIndex, int targetCellIndex) {
        log.info("=======dealMultiValueOneRowLogisticData==========");
        for (Integer valueOneRowIndex : multiValueOneRowIndex) {
            String cellValue = mainSheet.getRow(valueOneRowIndex).getCell(logisticColIndex).getStringCellValue();
            String[] array = cellValue.split("\n");
            BigDecimal totalValue = BigDecimal.ZERO;
            for (int i = 0; i < array.length; i++) {
                totalValue = totalValue.add(Convert.toBigDecimal(logisticsMap.getOrDefault(StrUtil.trim(array[i]), BigDecimal.ZERO)));
            }
            XSSFCell logisticsCell = mainSheet.getRow(valueOneRowIndex).createCell(targetCellIndex);
            logisticsCell.setCellValue(Convert.toDouble(totalValue));
        }
    }

    /**
     * @return java.util.List<org.apache.poi.ss.util.CellRangeAddress>
     * @Author chenec
     * @Date 2024/2/17 9:43
     * @Description 获取合并单元格的元素
     * @Param [mainSheet, cellIndex]
     **/
    private static List<CellRangeAddress> getMergeCellList(XSSFSheet mainSheet, Integer cellIndex) {
//        List<Integer> mergeCellIndex = new ArrayList<>();
        List<CellRangeAddress> mergeCellList = new ArrayList<>();
        for (int i = 0; i < mainSheet.getNumMergedRegions(); i++) {
            CellRangeAddress cellRangeAddress = mainSheet.getMergedRegion(i);
            if (cellRangeAddress.getFirstColumn() != cellIndex) {
                continue;
            }
//            IntStream.range(cellRangeAddress.getFirstRow(), cellRangeAddress.getLastRow() + 1).forEach(t -> mergeCellIndex.add(t));
            mergeCellList.add(cellRangeAddress);
        }
        return mergeCellList;
    }

    /**
     * @return java.util.List<java.lang.Integer>
     * @Author chenec
     * @Date 2024/2/17 9:35
     * @Description 获取一个单元格多个值的下标值集合
     * @Param [mainSheet]
     **/
    private static List<Integer> getMultiValueOneCellList(XSSFSheet mainSheet, Integer cellIndex, List<Integer> mergeCellIndex) {
        List<Integer> multiValueOneCellIndex = new ArrayList<>();
        for (int i = 1; i <= mainSheet.getLastRowNum(); i++) {
            XSSFCell cell = mainSheet.getRow(i).getCell(cellIndex);
            if (null == cell) {
                continue;
            }
            String stringCellValue = cell.getStringCellValue();
            if (stringCellValue.split("\n").length > 1 && !CollUtil.contains(mergeCellIndex, i)) {
                multiValueOneCellIndex.add(i);
            }
        }
        return multiValueOneCellIndex;
    }

    /**
     * @return java.util.Map<java.lang.String, java.math.BigDecimal>
     * @Author chenec
     * @Date 2024/2/17 14:41
     * @Description 获取商店集合
     * @Param [workbook, sheet4, keyIndex, valueIndex]
     **/
    private static Map<String, ShopInfo> getShopMap(XSSFWorkbook workbook, String sheetName, int orderIndex, int shopIndex, int priceIndex) {
        XSSFSheet shopSheet = workbook.getSheet(sheetName);
        Map<String, ShopInfo> shopMap = new HashMap<>();
        List<String> ignoreList = ListUtil.of("微信", "代理", "订单编号\n" + "(除了单号不要写别的)", "");
        for (int i = 1; i <= shopSheet.getLastRowNum(); i++) {
            XSSFRow shopSheetRow = shopSheet.getRow(i);
            ShopInfo shopInfo = new ShopInfo();
            if (CollUtil.contains(ignoreList, shopSheetRow.getCell(orderIndex).getStringCellValue())) {
                log.info("=======跳过特殊行=======");
                continue;
            }
            BigDecimal price = BigDecimal.ZERO;
            if (shopSheetRow.getCell(priceIndex).getCellTypeEnum().equals(CellType.STRING)) {
                String priceType = shopSheetRow.getCell(priceIndex).getStringCellValue();
                if ("退款".equals(priceType)) {
                    price = new BigDecimal("-1");
                } else if ("破损退回重发".equals(priceType)) {
                    price = new BigDecimal("-2");
                } else {
                    price = new BigDecimal("-3");
                }
            } else {
                price = Convert.toBigDecimal(shopSheetRow.getCell(priceIndex).getNumericCellValue());
            }
            shopInfo.setShopName(shopSheetRow.getCell(shopIndex).getStringCellValue());
            shopInfo.setPrice(price);
            shopInfo.setOrderNo(shopSheetRow.getCell(orderIndex).getStringCellValue());
//            log.info("{}", shopInfo.getOrderNo());
            shopMap.put(shopInfo.getOrderNo(), shopInfo);
        }
        List<CellRangeAddress> mergeCellList = getMergeCellList(shopSheet, shopIndex);
        for (CellRangeAddress cellRangeAddress : mergeCellList) {
            int firstRow = cellRangeAddress.getFirstRow();
            int lastRow = cellRangeAddress.getLastRow();
            String shopName = shopSheet.getRow(firstRow).getCell(shopIndex).getStringCellValue();
            for (int k = firstRow; k <= lastRow; k++) {
                String orderNo = shopSheet.getRow(k).getCell(orderIndex).getStringCellValue();
                if (StrUtil.isBlank(orderNo)) {
                    continue;
                }
                shopMap.get(orderNo).setShopName(shopName);
            }
        }
        log.info("================================:{}", shopMap.size());
        return shopMap;
    }

    /**
     * @return java.util.Map<java.lang.String, java.math.BigDecimal>
     * @Author chenec
     * @Date 2024/2/17 14:45
     * @Description 获取物流费用集合
     * @Param [workbook, sheetName]
     **/
    private static Map<String, BigDecimal> getLogisticsMap(XSSFWorkbook workbook, String sheetName) {
        XSSFSheet logisticsSheet = workbook.getSheet(sheetName);
        Map<String, BigDecimal> logisticsMap = new HashMap<>();
        for (int i = 1; i <= logisticsSheet.getLastRowNum(); i++) {
            XSSFRow logisticsSheetRow = logisticsSheet.getRow(i);
            logisticsMap.put(logisticsSheetRow.getCell(0).getStringCellValue(), Convert.toBigDecimal(logisticsSheetRow.getCell(1).getNumericCellValue()));
        }
        return logisticsMap;
    }
    /**
     * @Author chenec
     * @Date 2024/2/18 11:33
     * @Description TODO 获取物流跟金额列的集合
     * @Param [workbook, sheetName]
     * @return java.util.Map<java.lang.String,java.math.BigDecimal>
     **/
    private static Map<String, BigDecimal> getLogisticsFcyMap(XSSFWorkbook workbook, String sheetName,int logisticsIndex) {
        XSSFSheet logisticsSheet = workbook.getSheet(sheetName);
        Map<String, BigDecimal> logisticsMap = new HashMap<>();
        for (int i = 1; i <= logisticsSheet.getLastRowNum(); i++) {
            XSSFRow logisticsSheetRow = logisticsSheet.getRow(i);
            logisticsMap.put(logisticsSheetRow.getCell(0).getStringCellValue(), Convert.toBigDecimal(logisticsSheetRow.getCell(1).getNumericCellValue()));
        }
        return logisticsMap;
    }

    private static String getValue(Sheet sheet, int row, int column) {
        //获取合并单元格的总数，并循环每一个合并单元格，
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            //判断当前单元格是否在合并单元格区域内，是的话就是合并单元格
            if ((row >= firstRow && row <= lastRow) && (column >= firstColumn && column <= lastColumn)) {
                Row fRow = sheet.getRow(firstRow);
                Cell fCell = fRow.getCell(firstColumn);
                //获取合并单元格首格的值
                return fCell.getStringCellValue();
            }
        }
        //非合并单元格个返回空
        return "";
    }
}
