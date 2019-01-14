package cn.cuslik;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;

/**
 * @Author:zhangchundong
 * @Date:Create in 11:19 2018/12/25
 */

public class SxssfServlet extends HttpServlet {
   @Override
   protected void doGet(HttpServletRequest req, HttpServletResponse resp) throws ServletException, IOException {
   }

   @Override
   protected void doPost(HttpServletRequest req, HttpServletResponse resp) throws ServletException, IOException {
      Sxssf sxssf = new Sxssf();
      try {
         System.out.println("name = "+req.getParameter("name"));
         System.out.println("age = "+req.getParameter("age"));
         sxssf.buildExcelDocument(req,resp);
         System.out.println("这是版本2.0");
      } catch (Exception e) {
         e.printStackTrace();
      }
   }
}
