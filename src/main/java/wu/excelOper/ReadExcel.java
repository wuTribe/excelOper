package wu.excelOper;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

public class ReadExcel {
    // 对应后缀名
    private static final String EXCEL_XLS = "xls";
    private static final String EXCEL_XLSX = "xlsx";

    /**
     * 读取excel中的文件
     *
     * @param file 文件
     * @return 数据集合
     */
    public List<String[]> readExcel(File file) {
        try {
            // 创建输入流，读取Excel
            Workbook wb = getWorkbok(file);
            // 获取Excel的第一页 sheet
            Sheet sheet = wb.getSheetAt(0);
            // 第一页总数据
            List<String[]> outerList = new ArrayList<String[]>();
            // sheet.getRows()返回该页的总行数
            for (int i = 0; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i); // 获得行对象
                String[] innerList = new String[row.getLastCellNum()]; // 存储一行内的数据的容器

                // sheet.getColumns()返回该页的总列数，遍历一列
                for (int j = 0; j < row.getLastCellNum(); j++) {
                    innerList[j] = changeContent(row.getCell(j)); // 将结果存入容器
                }

                outerList.add(i, innerList);
            }

            return outerList;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    /**
     * 针对数字格式进行格式化
     *
     * @param cell 单元格
     * @return 格式后的字符串
     */
    private String changeContent(Cell cell) {

        // 如果为空
        if (cell == null) {
            return "";
        }

        // 如果是数字单元格
        if (cell.getCellType() == CellType.NUMERIC) {
            // 为日期 格式化时间
            if (HSSFDateUtil.isCellDateFormatted(cell)) {
                SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                return sdf.format(cell.getDateCellValue());
            }

            // 不为日期 有可能是ID，电话类 价格，将单元格变为字符格式
            cell.setCellType(CellType.STRING);
            return cell.getStringCellValue();
        }

        // 其他照常显示
        return cell.getStringCellValue();
    }

    /**
     * 判断Excel的版本,获取Workbook高级操作对象
     *
     * @param file 文件
     * @return 对应文件的版本
     * @throws IOException 异常
     */
    private static Workbook getWorkbok(File file) throws IOException {
        Workbook wb = null;
        FileInputStream in = new FileInputStream(file);

        if (file.getName().endsWith(EXCEL_XLS)) {     //Excel&nbsp;2003
            wb = new HSSFWorkbook(in);
        } else if (file.getName().endsWith(EXCEL_XLSX)) {    // Excel 2007/2010
            wb = new XSSFWorkbook(in);
        }
        return wb;
    }
}