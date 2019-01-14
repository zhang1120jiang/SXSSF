package cn.cuslik;

import com.sun.xml.internal.ws.client.RequestContext;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;

/**
 * @Author:zhangchundong
 * @Date:Create in 9:56 2018/12/25
 */

public class Sxssf {
   public void buildExcelDocument(HttpServletRequest request, HttpServletResponse response) throws Exception {
      try {
         SXSSFWorkbook workbook = new SXSSFWorkbook(100);
         response.reset();
         // 注意，如果去掉下面一行代码中的attachment; 那么也会使IE自动打开文件。
         response.setHeader(
                 "Content-Disposition",
                 "attachment; filename="
                         + java.net.URLEncoder.encode(
                         "测试" + ".xlsx", "UTF-8"));//Excel 扩展名指定为xlsx  SXSSFWorkbook对象只支持xlsx格式
         OutputStream os = response.getOutputStream();
         CellStyle style = workbook.createCellStyle();
         // 设置样式
         style.setFillForegroundColor(HSSFColor.SKY_BLUE.index);//设置单元格着色
         style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);  //设置单元格填充样式
         style.setBorderBottom(HSSFCellStyle.BORDER_THIN);//设置下边框
         style.setBorderLeft(HSSFCellStyle.BORDER_THIN);//设置左边框
         style.setBorderRight(HSSFCellStyle.BORDER_THIN);//设置右边框
         style.setBorderTop(HSSFCellStyle.BORDER_THIN);//上边框
         style.setAlignment(HSSFCellStyle.ALIGN_CENTER);// 居中
         //List<EmployeeTrainHistoryModel> list = null;
         try {
            //查询数据库中共有多少条数据
            // query.setTotalItem(employeeTrainHistoryService.fetchCountEmployeeTrainHistoryByQuery(query));

            int page_size = 100000;// 定义每页数据数量
            int list_count = 200000;
            //总数量除以每页显示条数等于页数
            int export_times = list_count % page_size > 0 ? list_count / page_size
                    + 1 : list_count / page_size;
            //循环获取产生每页数据
            for (int m = 0; m < export_times; m++) {
               // query.setNeedQueryAll(false);
               //  query.setPageSize(100000);//每页显示多少条数据
               //  query.setCurrentPage(m+1);//设置第几页
               // list=employeeTrainHistoryService.getEmployeeTrainHistoryByQuery(query);
               //新建sheet
               Sheet sheet = null;
               sheet = workbook.createSheet(System.currentTimeMillis() + "");
               // 创建属于上面Sheet的Row，参数0可以是0～65535之间的任何一个，
               Row header = sheet.createRow(0); // 第0行
               // 产生标题列，每个sheet页产生一个标题
               Cell cell;
               String[] headerArr = new String[]{"标题一", "标题二",
                       "标题三", "标题四", "标题五", "标题六", "标题七", "标题八",
                       "标题九"};
               for (int j = 0; j < headerArr.length; j++) {
                  cell = header.createCell((short) j);
                  cell.setCellStyle(style);
                  cell.setCellValue(headerArr[j]);
               }
               // 迭代数据
               if (/*list != null && list.size() > 0*/true) {
                  int rowNum = 1;
                  for (int i = 0; i < 1000000; i++) {
                     //EmployeeTrainHistoryModel history=list.get(i);
                     sheet.setDefaultColumnWidth((short) 17);
                     Row row = sheet.createRow(rowNum++);
                     row.createCell((short) 0).setCellValue("测试");
                     row.createCell((short) 1).setCellValue("测试");
                     row.createCell((short) 2).setCellValue("测试");
                     row.createCell((short) 3).setCellValue("测试");
                     row.createCell((short) 4).setCellValue("测试");
                     row.createCell((short) 5).setCellValue("测试");
                     row.createCell((short) 6).setCellValue("测试");
                     row.createCell((short) 7).setCellValue("测试");
                     row.createCell((short) 8).setCellValue("测试");
                  }
               }
               //list.clear();
            }
         } catch (Exception e) {
            e.printStackTrace();
         }
         try {
            workbook.write(os);
            os.close();
         } catch (Exception e) {
            e.printStackTrace();
         }
      } catch (IOException e) {
         e.printStackTrace();
      }
   }
}