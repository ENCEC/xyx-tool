package org.example.service;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.collection.ListUtil;
import cn.hutool.core.convert.Convert;
import cn.hutool.core.date.DatePattern;
import cn.hutool.core.date.DateUtil;
import cn.hutool.core.io.FileUtil;
import cn.hutool.core.util.ObjectUtil;
import cn.hutool.core.util.StrUtil;
import cn.hutool.poi.excel.cell.values.ErrorCellValue;
import cn.hutool.poi.excel.cell.values.NumericCellValue;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.example.ShopInfo;
import org.example.dto.FurnitureLogisticDto;
import org.example.dto.FurnitureSpecDto;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.BufferedOutputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.net.URLEncoder;
import java.util.*;
import java.util.function.Function;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

/**
 * @Auther: chenec
 * @Date: 2024/2/19 09:30
 * @Description: ShopService
 * @Version 1.0.0
 */
@Service
@Slf4j
public class ShopService {
    /**
     * @return void
     * @Author chenec
     * @Date 2024/2/20 16:11
     * @Description 生成屏风日报
     * @Param [multipartFile, response]
     **/
    public void generatorScreenDayReport(MultipartFile multipartFile, HttpServletResponse response) throws Exception {
        int logisticColIndex = 10;
        int orderColIndex = 12;
        int materialIndex = 7;
        int fanIndex = 5;
        int fanTouIndex = 16;
        //获取工作簿
        XSSFWorkbook workbook = new XSSFWorkbook(multipartFile.getInputStream());
        //获取工作表
        XSSFSheet mainSheet = workbook.getSheet("Sheet1");
        //update time 2024-07-12调整
        Map<String, FurnitureLogisticDto> orderMap = getScreenOrderMap(workbook, "Sheet2");
//        Map<String, BigDecimal> logisticsMap = getLogisticsMap(workbook, "Sheet4");

        Row firstRow = mainSheet.getRow(0);
        short lastCellNum = firstRow.getLastCellNum();
        //创建表头
//        String[] addCols = {"店铺", "成本", "物流", "订单号", "金额"};
        String[] addCols = {"店铺", "成本", "物流", "客服提成", "推广", "总成本", "金额", "扇数"};
        for (int i = 0; i < addCols.length; i++) {
            Cell cell = firstRow.createCell(lastCellNum + i);
            cell.setCellValue(addCols[i]);
        }
        //物流列合并单元格所占行下标集合
        List<Integer> mergeLogisticRowIndex = new ArrayList<>();
        List<CellRangeAddress> mergeLogisticCellAddressList = getMergeCellList(mainSheet, logisticColIndex);
        mergeLogisticCellAddressList.forEach(e -> {
            IntStream.range(e.getFirstRow(), e.getLastRow() + 1).forEach(t -> mergeLogisticRowIndex.add(t));
        });

//        List<Integer> multiLogisticValueOneRowIndex = getMultiValueOneCellList(mainSheet, logisticColIndex, mergeLogisticRowIndex);
//        dealMultiValueOneRowLogisticData(mainSheet, multiLogisticValueOneRowIndex, logisticsMap, logisticColIndex, lastCellNum + 2, true);
//        dealMergeRowLogisticData(mainSheet, mergeLogisticCellAddressList, logisticsMap, logisticColIndex, lastCellNum + 2, true);
        List<Integer> allRowIndexList = IntStream.range(1, mainSheet.getLastRowNum() + 1).boxed().collect(Collectors.toList());
//        Collection<Integer> subtractList1 = CollUtil.subtract(allRowIndexList, multiLogisticValueOneRowIndex);
//        Collection<Integer> normalLogisticIndexList = CollUtil.subtract(subtractList1, mergeLogisticRowIndex);
//        dealNormalRowLogisticData(mainSheet, normalLogisticIndexList, logisticsMap, logisticColIndex, lastCellNum + 2, true);
        dealLogisticsData(mainSheet, materialIndex, fanIndex, lastCellNum + 2);


        //店铺合并行下标集合
        List<Integer> mergeOrderRowIndex = new ArrayList<>();
        List<CellRangeAddress> mergeOrderCellAddressList = getMergeCellList(mainSheet, orderColIndex);
        mergeOrderCellAddressList.forEach(e -> {
            IntStream.range(e.getFirstRow(), e.getLastRow() + 1).forEach(t -> mergeOrderRowIndex.add(t));
        });
        //店铺普通列
        Collection<Integer> normalOrderIndexList = CollUtil.subtract(allRowIndexList, mergeOrderRowIndex);

        dealNormalRowOrderData(mainSheet, normalOrderIndexList, orderMap, orderColIndex, lastCellNum);
        dealMergeRowOrderData(mainSheet, mergeOrderCellAddressList, orderMap, orderColIndex, lastCellNum);
        dealOtherColOrderData(mainSheet, lastCellNum, fanIndex, fanTouIndex);

        dealAllCommonIndex(mainSheet, lastCellNum);

        response.setContentType("application/vnd.ms-excel;charset=UTF-8");
        response.setHeader("Content-Disposition", "attachment;filename=" + URLEncoder.encode(FileUtil.getPrefix(multipartFile.getOriginalFilename()), "UTF-8") + "-" + DateUtil.format(DateUtil.date(), DatePattern.PURE_DATETIME_FORMAT) + ".xlsx");
        response.setHeader("Pragma", "no-cache");
        response.setHeader("Cache-Control", "no-cache");
        response.setHeader("Expires", " 0");

        OutputStream output;
        try {
            output = response.getOutputStream();
            BufferedOutputStream bufferOutput = new BufferedOutputStream(output);
            bufferOutput.flush();
            workbook.write(bufferOutput);
            bufferOutput.close();
        } catch (Exception e) {
            log.error("exception:{}", e);
        }
    }

    /**
     * 日报处理物流列数据，根据材质及扇数计算物流信息
     *
     * @param mainSheet
     * @param materialIndex
     * @param fanIndex
     * @param targetCellIndex
     */
    private void dealLogisticsData(XSSFSheet mainSheet, int materialIndex, int fanIndex, int targetCellIndex) {
        int firstRowNum = mainSheet.getFirstRowNum();
        int lastRowNum = mainSheet.getLastRowNum();
        for (int i = firstRowNum + 1; i < lastRowNum + 1; i++) {
            String material = getCellValue(mainSheet.getRow(i).getCell(materialIndex));
            String fan = getCellValue(mainSheet.getRow(i).getCell(fanIndex));
            XSSFCell logisticsCell = mainSheet.getRow(i).createCell(targetCellIndex);
            if (StrUtil.contains(material,"座屏")) {
                BigDecimal logisticsValue = Convert.toBigDecimal(fan).multiply(new BigDecimal("25"));
                logisticsCell.setCellValue(logisticsValue.intValue());
            } else {
                if (Convert.toInt(fan) < 3) {
                    logisticsCell.setCellValue(15);
                } else {
                    BigDecimal logisticsValue = Convert.toBigDecimal(fan).multiply(new BigDecimal("5"));
                    logisticsCell.setCellValue(logisticsValue.intValue());
                }
            }
        }
    }

    private void dealAllCommonIndex(XSSFSheet mainSheet, short lastCellNum) {
        log.info("==========处理总成本和客服提成========");
        int custColIndex = lastCellNum + 3;
        int allCostIndex = lastCellNum + 5;
        int firstRowNum = mainSheet.getFirstRowNum();
        int lastRowNum = mainSheet.getLastRowNum();
        for (int i = firstRowNum + 1; i < lastRowNum + 1; i++) {
            log.info("row = {}", i);
            XSSFCell custCell = mainSheet.getRow(i).createCell(custColIndex);
            XSSFCell allCostCell = mainSheet.getRow(i).createCell(allCostIndex);
            XSSFCell logisticCell = mainSheet.getRow(i).getCell(custColIndex - 1);
            XSSFCell costCell = mainSheet.getRow(i).getCell(custColIndex - 2);
            XSSFCell fcyCell = mainSheet.getRow(i).getCell(allCostIndex + 1);
            //客服提成
            custCell.setCellValue(Convert.toBigDecimal(fcyCell.getRawValue()).multiply(new BigDecimal("0.01")).doubleValue());
            BigDecimal allCostBig = Convert.toBigDecimal(getCellValue(logisticCell)).add(Convert.toBigDecimal(getCellValue(costCell))).add(Convert.toBigDecimal(getCellValue(custCell)));
            allCostCell.setCellValue(allCostBig.doubleValue());
            if ("#".equals(getCellValue(fcyCell))) {
                custCell.setCellValue("#");
            }
            if ("#".equals(getCellValue(logisticCell)) || "#".equals(getCellValue(costCell)) || "#".equals(getCellValue(custCell))) {
                allCostCell.setCellValue("#");
            }
        }
    }

    private void dealOtherColOrderData(XSSFSheet mainSheet, short lastCellNum, int fanIndex, int fanTouIndex) {
        log.info("=============dealOtherColOrderData===========");
        for (int i = 1; i <= mainSheet.getLastRowNum(); i++) {
            XSSFRow sheetRow = mainSheet.getRow(i);
//            String employName = getCellValue(sheetRow.getCell(10));
//            //客服提成
//            BigDecimal fcyRate = new BigDecimal("0.01");
//            if ("馨".equals(employName) || "雅".equals(employName)) {
//                fcyRate = new BigDecimal("0.015");
//            }
//            XSSFCell commissionCell = sheetRow.createCell(lastCellNum + 3);
//            if ("#".equals(getCellValue(sheetRow.getCell(lastCellNum + 6)))) {
//                commissionCell.setCellValue("#");
//            } else {
//                commissionCell.setCellValue(Convert.toBigDecimal(getCellValue(sheetRow.getCell(lastCellNum + 6))).multiply(fcyRate).doubleValue());
//            }
            //成本
            log.info("i:{}", i);
            XSSFCell costCell = sheetRow.createCell(lastCellNum + 1);
            costCell.setCellValue(Convert.toBigDecimal(validateCellValueNum(sheetRow, fanTouIndex) ?
                    BigDecimal.ZERO : sheetRow.getCell(fanTouIndex).getRawValue()).add(
                    Convert.toBigDecimal(validateCellValueNum(sheetRow, fanTouIndex + 1) ?
                            BigDecimal.ZERO : sheetRow.getCell(fanTouIndex + 1).getRawValue()))
                    .add(Convert.toBigDecimal(validateCellValueNum(sheetRow, fanTouIndex + 2) ?
                            BigDecimal.ZERO : sheetRow.getCell(fanTouIndex + 2).getRawValue()))
                    .add(Convert.toBigDecimal(validateCellValueNum(sheetRow, fanTouIndex + 3) ?
                            BigDecimal.ZERO : sheetRow.getCell(fanTouIndex + 3).getRawValue()))
                    .add(Convert.toBigDecimal(validateCellValueNum(sheetRow, fanTouIndex + 4) ? BigDecimal.ZERO : sheetRow.getCell(fanTouIndex + 4).getRawValue()))
                    .add(Convert.toBigDecimal(validateCellValueNum(sheetRow, fanTouIndex + 5) ? BigDecimal.ZERO : sheetRow.getCell(fanTouIndex + 5).getRawValue()))
                    .doubleValue());
//            //总成本
//            XSSFCell totalCostCell = sheetRow.createCell(lastCellNum + 5);
//            totalCostCell.setCellValue(Convert.toBigDecimal("#".equals(getCellValue(sheetRow.getCell(lastCellNum + 1))) ? BigDecimal.ZERO : getCellValue(sheetRow.getCell(lastCellNum + 1)))
//                    .add("#".equals(getCellValue(sheetRow.getCell(lastCellNum + 2))) ? BigDecimal.ZERO : Convert.toBigDecimal(getCellValue(sheetRow.getCell(lastCellNum + 2))))
//                    .add("#".equals(getCellValue(sheetRow.getCell(lastCellNum + 3))) ? BigDecimal.ZERO : Convert.toBigDecimal(getCellValue(sheetRow.getCell(lastCellNum + 3))))
//                    .doubleValue());
            //扇数
            XSSFCell fanCell = sheetRow.createCell(lastCellNum + 7);
            fanCell.setCellValue(Convert.toInt(getCellValue(sheetRow.getCell(fanIndex))));
        }

    }

    private boolean validateCellValueNum(XSSFRow sheetRow, int cellColIndex) {
        if (null == sheetRow.getCell(cellColIndex)) {
            return true;
        }
        if (StrUtil.isBlank(sheetRow.getCell(cellColIndex).getRawValue())) {
            return true;
        }
        return false;
    }

    private void fillFurnitureSpecData(XSSFSheet mainSheet, Map<String, FurnitureSpecDto> specMap, int specIndex, int lastCellNum) {
        for (int i = 1; i <= mainSheet.getLastRowNum(); i++) {
            XSSFCell cell = mainSheet.getRow(i).getCell(specIndex);
            FurnitureSpecDto furnitureSpecDto = specMap.get(getCellValue(cell));
            if (null == furnitureSpecDto) {
                continue;
            }
            mainSheet.getRow(i).createCell(lastCellNum + 1).setCellValue(furnitureSpecDto.getProductNo());
            mainSheet.getRow(i).createCell(lastCellNum + 2).setCellValue(furnitureSpecDto.getCost().doubleValue());
        }
    }

    private void fillFurnitureLogisticData(XSSFSheet mainSheet, Map<String, FurnitureLogisticDto> logisticMap, int logisticColIndex, int lastCellNum) {
        for (int i = 1; i <= mainSheet.getLastRowNum(); i++) {
            XSSFCell cell = mainSheet.getRow(i).getCell(logisticColIndex);
            FurnitureLogisticDto furnitureLogisticDto = logisticMap.get(getCellValue(cell));
            if (null == furnitureLogisticDto) {
                continue;
            }
            mainSheet.getRow(i).createCell(lastCellNum).setCellValue(furnitureLogisticDto.getShopName());
            mainSheet.getRow(i).createCell(lastCellNum + 4).setCellValue(furnitureLogisticDto.getOrderNo());
            mainSheet.getRow(i).createCell(lastCellNum + 5).setCellValue(furnitureLogisticDto.getFcy().doubleValue());
        }
    }

    /**
     * @return void
     * @Author chenec
     * @Date 2024/2/20 16:12
     * @Description 生成屏风月报
     * @Param [multipartFile, response]
     **/
    public void generatorScreenMonthReport(MultipartFile multipartFile, HttpServletResponse response) throws Exception {
        int logisticColIndex = 9;
        int orderColIndex = 11;
        int[] costIndexArray = {15, 16, 17, 18, 19, 22};
        String[] wechatArray = {"微信", "代理"};
        String[] logisticSheetArray = {"顺心", "韵达", "安能", "德邦", "顺丰"};
        //获取工作簿
        XSSFWorkbook workbook = new XSSFWorkbook(multipartFile.getInputStream());
        //获取工作表
        XSSFSheet mainSheet = workbook.getSheet("总表");
        //获取物流数据
        Map<String, BigDecimal> logisticsMap = new HashMap<>();
        Map<String, BigDecimal> sxLogisticsMap = getLogisticsMap(workbook, "顺心");
        Map<String, BigDecimal> ydLogisticsMap = getLogisticsMap(workbook, "韵达");
        Map<String, BigDecimal> anLogisticsMap = getLogisticsMap(workbook, "安能");
        Map<String, BigDecimal> dbLogisticsMap = getLogisticsMap(workbook, "德邦");
        Map<String, BigDecimal> sfLogisticsMap = getSFLogisticsMap(workbook, "顺丰");
        logisticsMap.putAll(sxLogisticsMap);
        logisticsMap.putAll(ydLogisticsMap);
        logisticsMap.putAll(anLogisticsMap);
        logisticsMap.putAll(dbLogisticsMap);
        logisticsMap.putAll(sfLogisticsMap);
        log.info("获取物流数据总和：{}", logisticsMap.size());
        //获取店铺数据
//        Map<String, ShopInfo> shopMap = getShopMap(workbook, "Sheet4", 12, 13, 14);

        log.info("last row num: {}", mainSheet.getLastRowNum());
        log.info("last col num: {}", mainSheet.getRow(0).getLastCellNum());
        List<CellRangeAddress> mergeCellAddressList = getMergeCellList(mainSheet, logisticColIndex);
        //物流列合并单元格所占行下标集合
        List<Integer> mergeRowIndex = new ArrayList<>();
        mergeCellAddressList.forEach(e -> {
            IntStream.range(e.getFirstRow(), e.getLastRow() + 1).forEach(t -> mergeRowIndex.add(t));
        });

        List<Integer> multiValueOneRowIndex = getMultiValueOneCellList(mainSheet, logisticColIndex, mergeRowIndex);

        Row firstRow = mainSheet.getRow(0);
        short lastCellNum = firstRow.getLastCellNum();
        //创建表头
        String[] addCols = {"物流"};
        for (int i = 0; i < addCols.length; i++) {
            Cell cell = firstRow.createCell(lastCellNum + i);
            cell.setCellValue(addCols[i]);
        }
        dealMultiValueOneRowLogisticData(mainSheet, multiValueOneRowIndex, logisticsMap, logisticColIndex, lastCellNum, false);
        dealMergeRowLogisticData(mainSheet, mergeCellAddressList, logisticsMap, logisticColIndex, lastCellNum, false);
        List<Integer> allRowIndexList = IntStream.range(1, mainSheet.getLastRowNum() + 1).boxed().collect(Collectors.toList());
        Collection<Integer> subtractList1 = CollUtil.subtract(allRowIndexList, multiValueOneRowIndex);
        Collection<Integer> normalCollection = CollUtil.subtract(subtractList1, mergeRowIndex);
        dealNormalRowLogisticData(mainSheet, normalCollection, logisticsMap, logisticColIndex, lastCellNum, false);
        List<String> summaryLogisticList = getSummaryLogisticList(mainSheet, logisticColIndex);
        Map<String, String> logisticFanMap;
        try {
            logisticFanMap = getLogisticFanMap(mainSheet, logisticColIndex, normalCollection, multiValueOneRowIndex, mergeCellAddressList);
        } catch (Exception e) {
            throw new RuntimeException("解析扇数内容报错");
        }

        dealEveryLogisticSheet(workbook, summaryLogisticList, logisticSheetArray, normalCollection, multiValueOneRowIndex, mergeCellAddressList, logisticFanMap);


        response.setContentType("application/vnd.ms-excel;charset=UTF-8");
        response.setHeader("Content-Disposition", "attachment;filename=" + URLEncoder.encode(FileUtil.getPrefix(multipartFile.getOriginalFilename()), "UTF-8") + "-" + DateUtil.format(DateUtil.date(), DatePattern.PURE_DATETIME_FORMAT) + ".xlsx");
        response.setHeader("Pragma", "no-cache");
        response.setHeader("Cache-Control", "no-cache");
        response.setHeader("Expires", " 0");

        OutputStream output;
        try {
            output = response.getOutputStream();
            BufferedOutputStream bufferOutput = new BufferedOutputStream(output);
            bufferOutput.flush();
            workbook.write(bufferOutput);
            bufferOutput.close();
        } catch (Exception e) {
            log.error("exception:{}", e);
        }
    }

    /**
     * @return java.util.List<java.lang.String>
     * @Author chenec
     * @Date 2024/2/24 11:05
     * @Description 获取总表的所有物流列数据
     * @Param [mainSheet, logisticColIndex]
     **/
    private List<String> getSummaryLogisticList(XSSFSheet mainSheet, int logisticColIndex) {
        List<String> logisticList = new ArrayList<>();
        for (int i = 1; i <= mainSheet.getLastRowNum(); i++) {
            XSSFRow logisticsSheetRow = mainSheet.getRow(i);
            String cellValue = getCellValue(logisticsSheetRow, logisticColIndex);
            String[] array = cellValue.split("\n");
            if (array.length > 1) {
                for (String val : array) {
                    logisticList.add(val.trim());
                }
            } else {
                logisticList.add(cellValue.trim());
            }
        }
        return logisticList;
    }

    /**
     * @return java.util.Map<java.lang.String, java.lang.String> 返回<物流单号，扇数>
     * @Author chenec
     * @Date 2024/5/19 16:46
     * @Description
     * @Param [mainSheet, logisticColIndex, normalCollection, multiValueOneRowIndex, mergeCellAddressList]
     **/
    private Map<String, String> getLogisticFanMap(XSSFSheet mainSheet, int logisticColIndex, Collection<Integer> normalCollection, List<Integer> multiValueOneRowIndex, List<CellRangeAddress> mergeCellAddressList) {
        int fanNum = 4;
        Map<String, String> logisticFanMap = new HashMap<>();
        normalCollection.stream().forEach(e ->
                logisticFanMap.put(getCellValue(mainSheet.getRow(e), logisticColIndex).trim(), getCellValue(mainSheet.getRow(e), fanNum)));
        for (int i = 0; i < multiValueOneRowIndex.size(); i++) {
            String logisticNumStr = getCellValue(mainSheet.getRow(multiValueOneRowIndex.get(i)), logisticColIndex).trim();
            String[] splitLogistic = logisticNumStr.split("\n");
            for (int j = 0; j < splitLogistic.length; j++) {
                if (StrUtil.isBlank(splitLogistic[j].trim())) {
                    continue;
                }
                if (j == 0) {
                    logisticFanMap.put(splitLogistic[j].trim(), getCellValue(mainSheet.getRow(multiValueOneRowIndex.get(i)), fanNum));
                } else {
                    logisticFanMap.put(splitLogistic[j].trim(), "0");
                }
            }
        }
        for (CellRangeAddress cellRangeAddress : mergeCellAddressList) {
            int firstRow = cellRangeAddress.getFirstRow();
            int lastRow = cellRangeAddress.getLastRow();
            String logisticNum = getCellValue(mainSheet.getRow(firstRow), logisticColIndex);
            int logisticFanNum = 0;
            for (int i = firstRow; i <= lastRow; i++) {
                if (StrUtil.isBlank(getCellValue(mainSheet.getRow(i), fanNum))) {
                    continue;
                }
                logisticFanNum += Convert.toInt(getCellValue(mainSheet.getRow(i), fanNum));
            }
            logisticFanMap.put(logisticNum.split("\n")[0].trim(), String.valueOf(logisticFanNum));
        }
        return logisticFanMap;
    }

    /**
     * @return void
     * @Author chenec
     * @Date 2024/2/24 11:12
     * @Description 处理每个物流sheet页，没有匹配到的单号，添加一列标识
     * @Param [workbook, summaryLogisticList, logisticSheetArray]
     **/
    private void dealEveryLogisticSheet(XSSFWorkbook workbook, List<String> summaryLogisticList, String[] logisticSheetArray,
                                        Collection<Integer> normalCollection, List<Integer> multiValueOneRowIndex, List<CellRangeAddress> mergeCellAddressList, Map<String, String> logisticFanMap) {
        log.info("=========处理每个物流sheet页，没有匹配到的单号，添加一列标识==========");
        for (String sheetName : logisticSheetArray) {
            XSSFSheet sheet = workbook.getSheet(sheetName);
            log.info("处理sheet数据，sheet名称为：{}", sheetName);
            if (null == sheet) {
                log.info("跳过sheet，名称为：{}", sheetName);
                continue;
            }
            Row firstRow = sheet.getRow(0);
            if (null == firstRow) {
                throw new RuntimeException("物流企业提供的数据，第一行不能为空，请补充物流，金额两列数据。");
            }
            short lastCellNum = firstRow.getLastCellNum();
            XSSFCell firstCell = sheet.getRow(0).createCell(lastCellNum);
            XSSFCell secondCell = sheet.getRow(0).createCell(lastCellNum + 1);
            firstCell.setCellValue("是否匹配");
            secondCell.setCellValue("扇数");
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                String cellValue = getCellValue(sheet.getRow(i).getCell(0));
                if (StrUtil.isBlank(cellValue)) {
                    continue;
                }
                boolean contains = summaryLogisticList.contains(cellValue);
                XSSFCell cell = sheet.getRow(i).createCell(lastCellNum);
                if (contains) {
                    cell.setCellValue("是");
                } else {
                    cell.setCellValue("否");
                }
                XSSFCell secondCell1 = sheet.getRow(i).createCell(lastCellNum + 1);
                if (ObjectUtil.isNotNull(logisticFanMap.get(cellValue.trim()))) {
                    secondCell1.setCellValue(logisticFanMap.get(cellValue));
                } else {
                    secondCell1.setCellValue("0");
                }
            }
        }
    }


    /**
     * @return void
     * @Author chenec
     * @Date 2024/2/17 17:12
     * @Description 处理合并的订单列
     * @Param [mainSheet, normalShopIndexData, shopMap, shopColIndex, lastCellNum]
     **/
    private static void fillNormalOrderData(XSSFSheet mainSheet, Collection<Integer> normalShopIndexData, Map<String, ShopInfo> shopMap, int orderColIndex, int shopNameCellIndex, int fcyIndex, String[] wechatArray, int logisticColIndex) {
        log.info("=======fillNormalOrderData==========");
        for (Integer rowIndex : normalShopIndexData) {
            XSSFCell cell = mainSheet.getRow(rowIndex).getCell(orderColIndex);
            if (null == cell) {
                continue;
            }
            String cellValue = StrUtil.trim(getCellValue(cell));
            XSSFCell shopNameCell = mainSheet.getRow(rowIndex).createCell(shopNameCellIndex);
            XSSFCell fcyCell = mainSheet.getRow(rowIndex).createCell(fcyIndex);
            ShopInfo shopInfo = shopMap.get(cellValue);
            String shopName = ObjectUtil.isNull(shopInfo) ? "" : shopInfo.getShopName();
            BigDecimal price = ObjectUtil.isNull(shopInfo) ? BigDecimal.ZERO : shopInfo.getPrice();
            shopNameCell.setCellValue(shopName);
            fcyCell.setCellValue(price.doubleValue());
            if (shopInfo.getPrice().compareTo(BigDecimal.ZERO) < 0) {
                fcyCell.setCellValue("N/A");
            }
            if (Arrays.stream(wechatArray).anyMatch(e -> e.equals(cellValue))) {
                cell = mainSheet.getRow(rowIndex).getCell(logisticColIndex);

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
    private static void fillMergeOrderData(XSSFSheet mainSheet, List<CellRangeAddress> shopMergeCellAddressList, Map<String, ShopInfo> shopMap, int orderColIndex, int shopNameCellIndex, int fcyIndex, String[] wechatArray, int logisticColIndex) {
        log.info("=======fillNormalOrderData==========");
        for (CellRangeAddress cellRangeAddress : shopMergeCellAddressList) {
            int firstRow = cellRangeAddress.getFirstRow();
            int lastRow = cellRangeAddress.getLastRow();
            String cellValue = getCellValue(mainSheet.getRow(firstRow).getCell(orderColIndex));
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
    private static void dealNormalRowLogisticData(XSSFSheet mainSheet, Collection<Integer> normalCollection, Map<String, BigDecimal> logisticsMap, int logisticColIndex, int targetCellIndex, boolean dayFlag) {
        log.info("=======dealNormalRowLogisticData==========");
        for (Integer rowIndex : normalCollection) {
            XSSFCell cell = mainSheet.getRow(rowIndex).getCell(logisticColIndex);
            if (null == cell) {
                continue;
            }
            String cellValue = StrUtil.trim(getCellValue(cell));
            XSSFCell logisticsCell = mainSheet.getRow(rowIndex).createCell(targetCellIndex);
            if (null == logisticsMap.get(cellValue)) {
                logisticsCell.setCellValue("#");
            } else {
                logisticsCell.setCellValue(Convert.toDouble(logisticsMap.get(cellValue)));
            }
            //日报标识，物流列只查找S6开头的数据，其他的物流单号，根据扇数判断，如果扇数为1，则直接=20，不然扇数*7.
            if (dayFlag) {
//                if (!cellValue.startsWith("S6")) {
                if (null == logisticsMap.get(cellValue)) {
                    log.info("rowIndex:{}", rowIndex);
                    int num = Convert.toInt(getCellValue(mainSheet.getRow(rowIndex).getCell(4)));
                    logisticsCell.setCellValue(num > 1 ? num * 7 : 20);
                }
            }
        }
    }

    /**
     * @return void
     * @Author chenec
     * @Date 2024/2/25 14:16
     * @Description 处理普通行店铺数据
     * @Param [mainSheet, normalShopIndexList, shopMap, shopColIndex, lastCellNum]
     **/
    private void dealNormalRowOrderData(XSSFSheet mainSheet, Collection<Integer> normalOrderIndexList, Map<String, FurnitureLogisticDto> orderMap, int orderColIndex, short lastCellNum) {
        log.info("=======dealNormalRowOrderData==========");
        for (Integer rowIndex : normalOrderIndexList) {
            XSSFCell cell = mainSheet.getRow(rowIndex).getCell(orderColIndex);
            if (null == cell) {
                continue;
            }
            String cellValue = StrUtil.trim(getCellValue(cell));
            XSSFCell shopCell = mainSheet.getRow(rowIndex).createCell(lastCellNum);
            XSSFCell fcyCell = mainSheet.getRow(rowIndex).createCell(lastCellNum + 6);
            if (null == orderMap.get(cellValue)) {
                shopCell.setCellValue("#");
                fcyCell.setCellValue("#");
            } else {
                shopCell.setCellValue(orderMap.get(cellValue).getShopName());
                if (null == orderMap.get(cellValue).getFcy()) {
                    fcyCell.setCellValue("#");
                } else {
                    fcyCell.setCellValue(orderMap.get(cellValue).getFcy().doubleValue());
                }
            }
        }
    }

    /**
     * @return void
     * @Author chenec
     * @Date 2024/2/17 10:34
     * @Description 处理合并单元格的情况，物流数据查找填充
     * @Param [mainSheet, mergeRowIndex, logisticsMap, logisticColIndex, i]
     **/
    private static void dealMergeRowLogisticData(XSSFSheet mainSheet, List<CellRangeAddress> mergeCellAddressList, Map<String, BigDecimal> logisticsMap, int logisticColIndex, int targetCellIndex, boolean dayFlag) {
        log.info("=======dealMergeRowLogisticData==========");
        for (CellRangeAddress cellRangeAddress : mergeCellAddressList) {
            int firstRow = cellRangeAddress.getFirstRow();
            int lastRow = cellRangeAddress.getLastRow();
            String cellValue = StrUtil.trim(getCellValue(mainSheet.getRow(firstRow).getCell(logisticColIndex)));
            String[] array = cellValue.split("\n");
            BigDecimal totalValue = BigDecimal.ZERO;
            //不存在子母单号的物流，匹配标识
            boolean misMatchFlag = false;
            for (int i = 0; i < array.length; i++) {
                String cellVale = StrUtil.trim(array[i]);
//                if (!cellVale.startsWith("DPK") && StrUtil.isNotBlank(cellVale) && null == logisticsMap.get(cellVale)) {
                if (StrUtil.isNotBlank(cellVale) && null == logisticsMap.get(cellVale)) {
                    misMatchFlag = true;
                    break;
                }
                totalValue = totalValue.add(Convert.toBigDecimal(logisticsMap.getOrDefault(cellVale, BigDecimal.ZERO)));
            }
            //德邦匹配标识，匹配上一个就算通过
            for (int i = 0; i < array.length; i++) {
                String cellVale = StrUtil.trim(array[i]);
                if (cellVale.startsWith("DPK") && ObjectUtil.isNotNull(logisticsMap.get(cellVale))) {
                    misMatchFlag = false;
                    break;
                }
            }
            for (int i = firstRow; i <= lastRow; i++) {
                XSSFCell logisticsCell = mainSheet.getRow(i).createCell(targetCellIndex);
                if (misMatchFlag) {
                    logisticsCell.setCellValue("#");
                } else {
                    logisticsCell.setCellValue(0L);
                }
            }
            XSSFCell logisticsCell1 = mainSheet.getRow(firstRow).createCell(targetCellIndex);
            if (misMatchFlag) {
                logisticsCell1.setCellValue("#");
            } else {
                logisticsCell1.setCellValue(Convert.toDouble(totalValue));
            }
            //日报标识，物流列只查找S6开头的数据，其他的物流单号，根据扇数判断，如果扇数为1，则直接=20，不然扇数*7.
            //取第一个判断是否是顺心的，如果单号组合为顺心+德邦在同一个合并单元格，则计算值有误
            String cellVale = StrUtil.trim(array[0]);
//            if (dayFlag && !cellVale.startsWith("S6")) {
            if (dayFlag && misMatchFlag) {
                XSSFCell logisticsCell;
                int num = 0;
                for (int i = firstRow; i <= lastRow; i++) {
                    logisticsCell = mainSheet.getRow(i).createCell(targetCellIndex);
//                    if (!cellVale.startsWith("S6")) {
                    logisticsCell.setCellValue(0L);
                    if (StrUtil.isBlank(getCellValue(mainSheet.getRow(i).getCell(4)))) {
                        continue;
                    }
                    num = num + Convert.toInt(getCellValue(mainSheet.getRow(i).getCell(4)));
//                    }
                }
                logisticsCell = mainSheet.getRow(firstRow).createCell(targetCellIndex);
                logisticsCell.setCellValue(num > 1 ? num * 7 : 20);
            }
        }
    }

    /**
     * @return void
     * @Author chenec
     * @Date 2024/2/25 15:06
     * @Description 处理合并单元格的情况，店铺及金额数据查找填充
     * @Param [mainSheet, normalOrderIndexList, orderMap, orderColIndex, lastCellNum]
     **/

    private void dealMergeRowOrderData(XSSFSheet mainSheet, List<CellRangeAddress> mergeOrderCellAddressList, Map<String, FurnitureLogisticDto> orderMap, int orderColIndex, short targetCellIndex) {
        log.info("=======dealMergeRowOrderData==========");
        for (CellRangeAddress cellRangeAddress : mergeOrderCellAddressList) {
            int firstRow = cellRangeAddress.getFirstRow();
            int lastRow = cellRangeAddress.getLastRow();
            String cellValue = StrUtil.trim(getCellValue(mainSheet.getRow(firstRow).getCell(orderColIndex)));
            boolean misMatchFlag = false;
            if (null == orderMap.get(cellValue)) {
                misMatchFlag = true;
            }
            BigDecimal totalValue = Convert.toBigDecimal(null == orderMap.get(cellValue) || null == orderMap.get(cellValue).getFcy() ? BigDecimal.ZERO : orderMap.get(cellValue).getFcy());
            for (int i = firstRow; i <= lastRow; i++) {
                XSSFCell shopCell = mainSheet.getRow(i).createCell(targetCellIndex);
                XSSFCell fcyCell = mainSheet.getRow(i).createCell(targetCellIndex + 6);
                if (misMatchFlag) {
                    shopCell.setCellValue("#");
                    fcyCell.setCellValue("#");
                } else {
                    shopCell.setCellValue(orderMap.get(cellValue).getShopName());
                    fcyCell.setCellValue(0L);
                }
            }
            XSSFCell fcy1 = mainSheet.getRow(firstRow).createCell(targetCellIndex + 6);
            if (misMatchFlag) {
                fcy1.setCellValue("#");
            } else {
                fcy1.setCellValue(Convert.toDouble(totalValue));
            }
        }
    }

    /**
     * @return void
     * @Author chenec
     * @Date 2024/2/17 10:29
     * @Description 处理一个单元格内多值的情况，物流数据查找填充
     * @Param [mainSheet, multiValueOneRowIndex, logisticsMap, logisticColIndex, targetCellIndex]
     **/
    private static void dealMultiValueOneRowLogisticData(XSSFSheet mainSheet, List<Integer> multiValueOneRowIndex, Map<String, BigDecimal> logisticsMap, int logisticColIndex, int targetCellIndex, boolean dayFlag) {
        log.info("=======dealMultiValueOneRowLogisticData==========");
        for (Integer valueOneRowIndex : multiValueOneRowIndex) {
            String cellValue = getCellValue(mainSheet.getRow(valueOneRowIndex).getCell(logisticColIndex));
            String[] array = cellValue.split("\n");
            BigDecimal totalValue = BigDecimal.ZERO;
            //不存在子母单号的物流，匹配标识
            boolean misMatchFlag = false;
            for (int i = 0; i < array.length; i++) {
                String cellVale = StrUtil.trim(array[i]);
                if (!cellVale.startsWith("DPK") && StrUtil.isNotBlank(cellVale) && null == logisticsMap.get(cellVale)) {
                    misMatchFlag = true;
                    break;
                }
                totalValue = totalValue.add(Convert.toBigDecimal(logisticsMap.getOrDefault(cellVale, BigDecimal.ZERO)));
            }
            //德邦匹配，匹配上一个就算通过
            for (int i = 0; i < array.length; i++) {
                String cellVale = StrUtil.trim(array[i]);
                if (cellVale.startsWith("DPK") && ObjectUtil.isNotNull(logisticsMap.get(cellVale))) {
                    misMatchFlag = false;
                    break;
                }
            }
            XSSFCell logisticsCell = mainSheet.getRow(valueOneRowIndex).createCell(targetCellIndex);
            if (misMatchFlag) {
                logisticsCell.setCellValue("#");
            } else {
                logisticsCell.setCellValue(Convert.toDouble(totalValue));
            }
            //日报标识，物流列只查找S6开头的数据，其他的物流单号，根据扇数判断，如果扇数为1，则直接=20，不然扇数*7.
            if (dayFlag) {
//                if (!StrUtil.trim(cellValue).startsWith("S6")) {
                int num = Convert.toInt(getCellValue(mainSheet.getRow(valueOneRowIndex).getCell(4)));
                logisticsCell.setCellValue(num > 1 ? num * 7 : 20);
//                }
            }
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
        log.info("====处理合并单元格数据====");
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
        log.info("====获取一个单元格多个值的下标值集合====");
        List<Integer> multiValueOneCellIndex = new ArrayList<>();
        for (int i = 1; i <= mainSheet.getLastRowNum(); i++) {
            XSSFCell cell = mainSheet.getRow(i).getCell(cellIndex);
            if (null == cell) {
                continue;
            }
            String stringCellValue = getCellValue(cell);
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
     * @Date 2024/2/18 14:22
     * @Description 获取顺丰物流数据
     * @Param [workbook, sheetName]
     **/
    private static Map<String, BigDecimal> getSFLogisticsMap(XSSFWorkbook workbook, String sheetName) {
        XSSFSheet logisticsSheet = workbook.getSheet(sheetName);
        Map<String, BigDecimal> logisticsMap = new HashMap<>();
        if (null == logisticsSheet) {
            return Collections.emptyMap();
        }
        for (int i = 1; i <= logisticsSheet.getLastRowNum(); i++) {
            XSSFRow logisticsSheetRow = logisticsSheet.getRow(i);
            String key = getCellValue(logisticsSheetRow, 0);
            BigDecimal value = Convert.toBigDecimal(getCellValue(logisticsSheetRow, 1));
            if (ObjectUtil.isEmpty(logisticsMap.get(key))) {
                logisticsMap.put(key, value);
            } else {
                logisticsMap.put(key, value.add(logisticsMap.get(key)));
            }
        }
        return logisticsMap;
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
        if (null == logisticsSheet) {
            throw new RuntimeException(sheetName + "sheet页不存在。");
        }
        log.info("=====获取{}物流数据，总行数为：{}。====", sheetName, logisticsSheet.getLastRowNum());
        Map<String, BigDecimal> logisticsMap = new HashMap<>();
        if (null == logisticsSheet) {
            return Collections.emptyMap();
        }
        for (int i = 1; i <= logisticsSheet.getLastRowNum(); i++) {
            XSSFRow logisticsSheetRow = logisticsSheet.getRow(i);
            if (StrUtil.isBlank(getCellValue(logisticsSheetRow, 1))) {
                throw new RuntimeException(String.format("读取表格发生异常，请检查表格。sheetName:%s,行：%s，列：%s。", sheetName, i + 1, 2));
            }
            logisticsMap.put(getCellValue(logisticsSheetRow, 0), Convert.toBigDecimal(getCellValue(logisticsSheetRow, 1)));
        }
        return logisticsMap;
    }

    private static String getCellValue(XSSFRow row, int cellIndex) {
        if (null == row || null == row.getCell(cellIndex)) {
            return "";
        }
        CellType cellTypeEnum = row.getCell(cellIndex).getCellTypeEnum();
        XSSFCell cell = row.getCell(cellIndex);
        Object value;
        switch (cellTypeEnum) {
            case NUMERIC:
                value = (new NumericCellValue(cell)).getValue();
                break;
            case BOOLEAN:
                value = cell.getBooleanCellValue();
                break;
            case BLANK:
                value = "";
                break;
            case FORMULA:
                value = cell.getRawValue();
                break;
            case ERROR:
                value = (new ErrorCellValue(cell)).getValue();
                break;
            default:
                value = cell.getStringCellValue();
        }

        return Convert.toStr(value);
    }

    private static String getCellValue(XSSFCell cell) {
        if (null == cell) {
            return "";
        }
        CellType cellTypeEnum = cell.getCellTypeEnum();
        Object value;
        switch (cellTypeEnum) {
            case NUMERIC:
                value = (new NumericCellValue(cell)).getValue();
                break;
            case BOOLEAN:
                value = cell.getBooleanCellValue();
                break;
            case BLANK:
                value = "";
                break;
            case ERROR:
                value = (new ErrorCellValue(cell)).getValue();
                break;
            case FORMULA:
                value = "#";
                break;
            default:
                value = cell.getStringCellValue();
        }
        return Convert.toStr(value);
    }

    /**
     * @return java.util.Map<java.lang.String, java.math.BigDecimal>
     * @Author chenec
     * @Date 2024/2/18 11:33
     * @Description TODO 获取物流跟金额列的集合
     * @Param [workbook, sheetName]
     **/
    private static Map<String, BigDecimal> getLogisticsFcyMap(XSSFWorkbook workbook, String sheetName, int logisticsIndex) {
        XSSFSheet logisticsSheet = workbook.getSheet(sheetName);
        Map<String, BigDecimal> logisticsMap = new HashMap<>();
        for (int i = 1; i <= logisticsSheet.getLastRowNum(); i++) {
            XSSFRow logisticsSheetRow = logisticsSheet.getRow(i);
            logisticsMap.put(getCellValue(logisticsSheetRow, 0), Convert.toBigDecimal(getCellValue(logisticsSheetRow, 1)));
        }
        return logisticsMap;
    }

    /**
     * @return java.util.Map<java.lang.String, org.example.dto.FurnitureSpecDto>
     * @Author chenec
     * @Date 2024/2/20 16:50
     * @Description 读取家居规格Sheet内容
     * @Param [workbook, sheetName]
     **/
    private Map<String, FurnitureSpecDto> getFurnitureSpecMap(XSSFWorkbook workbook, String sheetName) {
        List<FurnitureSpecDto> furnitureSpecDtos = new ArrayList<>();
        XSSFSheet specSheet = workbook.getSheet(sheetName);
        if (null == specSheet) {
            return Collections.emptyMap();
        }
        for (int i = 1; i <= specSheet.getLastRowNum(); i++) {
            XSSFRow specRow = specSheet.getRow(i);
            FurnitureSpecDto furnitureSpecDto = new FurnitureSpecDto();
            furnitureSpecDto.setSpec(getCellValue(specRow, 3));
            furnitureSpecDto.setProductNo(getCellValue(specRow, 4));
            furnitureSpecDto.setCost(Convert.toBigDecimal(getCellValue(specRow, 6)));
            furnitureSpecDtos.add(furnitureSpecDto);
        }
        return furnitureSpecDtos.stream().collect(Collectors.toMap(FurnitureSpecDto::getSpec, Function.identity(), (key1, key2) -> key2));
    }

    /**
     * @return java.util.Map<java.lang.String, org.example.dto.FurnitureLogisticDto>
     * @Author chenec
     * @Date 2024/2/20 16:49
     * @Description 读取屏风Sheet4店铺金额内容
     * @Param [workbook, sheetName]
     **/
    private Map<String, FurnitureLogisticDto> getScreenOrderMap(XSSFWorkbook workbook, String sheetName) {
        int logisticColIndex = 9;
        int orderIndex = 11;
        int shopIndex = 12;
        int fcyIndex = 13;
        List<FurnitureLogisticDto> furnitureLogisticDtos = new ArrayList<>();
        XSSFSheet shopSheet = workbook.getSheet(sheetName);
        if (null == shopSheet) {
            return Collections.emptyMap();
        }

        List<CellRangeAddress> mergeShopCellAddressList = getMergeCellList(shopSheet, shopIndex);

        for (int i = 1; i <= shopSheet.getLastRowNum(); i++) {
            XSSFRow logisticsSheetRow = shopSheet.getRow(i);
            String orderNo = getCellValue(logisticsSheetRow, orderIndex);
            if (StrUtil.isBlank(orderNo) || orderNo.startsWith("订单编号") || orderNo.equals("微信")) {
                continue;
            }
            FurnitureLogisticDto furnitureLogisticDto = new FurnitureLogisticDto();
            furnitureLogisticDto.setLogisticNo(getCellValue(logisticsSheetRow, logisticColIndex));
            furnitureLogisticDto.setOrderNo(orderNo);
            furnitureLogisticDto.setFcy(Convert.toBigDecimal(getCellValue(logisticsSheetRow, fcyIndex)));
            String shopName = getCellValue(logisticsSheetRow, shopIndex);
            for (CellRangeAddress cellRangeAddress : mergeShopCellAddressList) {
                int firstRow = cellRangeAddress.getFirstRow();
                int lastRow = cellRangeAddress.getLastRow();
                if (firstRow != i && lastRow != i) {
                    continue;
                }
                shopName = shopSheet.getRow(firstRow).getCell(shopIndex).getStringCellValue();
            }
            furnitureLogisticDto.setShopName(shopName);
            furnitureLogisticDtos.add(furnitureLogisticDto);
        }
        return furnitureLogisticDtos.stream().collect(Collectors.toMap(FurnitureLogisticDto::getOrderNo, Function.identity(), (key1, key2) -> key2));
    }
}
