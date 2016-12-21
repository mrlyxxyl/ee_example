package controller;

import net.sf.jxls.transformer.XLSTransformer;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;

import javax.servlet.ServletException;
import javax.servlet.ServletOutputStream;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.BufferedInputStream;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@WebServlet(urlPatterns = {"/servlet/excel"})
public class ExcelServlet extends HttpServlet {
    public void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
        Map<String, Object> map = new HashMap<String, Object>();
        List<Map<String, String>> users = new ArrayList<Map<String, String>>();
        Map<String, String> map1 = new HashMap<String, String>();
        map1.put("id", "1");
        map1.put("name", "jack");
        map1.put("age", "25");

        Map<String, String> map2 = new HashMap<String, String>();
        map2.put("id", "2");
        map2.put("name", "tom");
        map2.put("age", "26");

        users.add(map1);
        users.add(map2);
        map.put("users", users);

        String templateFileName = this.getServletContext().getRealPath("/doc/tmp.xls");
        String resultFileName = System.currentTimeMillis() + ".xls";

        ServletOutputStream os = null;
        InputStream is = null;
        try {
            os = response.getOutputStream();
            is = new BufferedInputStream(new FileInputStream(templateFileName));
            XLSTransformer transformer = new XLSTransformer();
            Workbook wb = transformer.transformXLS(is, map);
            response.reset();
            response.setHeader("Content-disposition", "attachment; filename=" + resultFileName);
            response.setContentType("application/msexcel");
            wb.write(os);
            is.close();
            os.close();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        } finally {
            if (os != null) {
                os.flush();
                os.close();
            }
            if (is != null) {
                is.close();
            }
        }
    }

    public void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
        doGet(request, response);
    }
}
