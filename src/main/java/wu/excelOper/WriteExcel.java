package wu.excelOper;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

public class WriteExcel {
    // 对应后缀名
    private static final String EXCEL_XLS = "xls";
    private static final String EXCEL_XLSX = "xlsx";

    /**
     * @param dataList      数据
     * @param finalXlsxPath 文件存放路径
     */
    public void writeExcel(List<String[]> dataList, String finalXlsxPath) {
        OutputStream out = null;
        try {
            // 创建文件
            File finalXlsxFile = new File(finalXlsxPath);
            // 生成对应版本的excel文件
            Workbook workBook = getWorkbok(finalXlsxFile);
            // sheet 对应一个工作页
            Sheet sheet = workBook.createSheet();
            // 创建文件输出流，输出电子表格：这个必须有，否则你在sheet上做的任何操作都不会有效
            out = new FileOutputStream(finalXlsxPath);

            // 往Excel中写新数据
            for (int j = 0; j < dataList.size(); j++) {
                Row row = sheet.createRow(j);

                // 在一行内循环
                for (int k = 0; k < dataList.get(j).length; ) {
                    for (Object o : dataList.get(j)) {
                        row.createCell(k).setCellValue((String) o);
                        k++;
                    }
                }
            }

            // 准备输出电子表格：这个必须有，否则你在sheet上做的任何操作都不会有效
            workBook.write(out);
            out.flush();
            System.out.println("\t\t数据导出成功");
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (out != null) {
                    out.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
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

        if (file.getName().endsWith(EXCEL_XLS)) {     //Excel&nbsp;2003
            wb = new HSSFWorkbook();
        } else if (file.getName().endsWith(EXCEL_XLSX)) {    // Excel 2007/2010
            wb = new XSSFWorkbook();
        }
        return wb;
    }
}