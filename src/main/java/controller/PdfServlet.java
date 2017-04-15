package controller;

import com.itextpdf.text.*;
import com.itextpdf.text.pdf.*;

import javax.servlet.ServletException;
import javax.servlet.ServletOutputStream;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@WebServlet(urlPatterns = {"/servlet/PdfServlet"})
public class PdfServlet extends HttpServlet {

    public static final String BASE_FONT = "c:/windows/fonts/simsun.ttc,1";

    protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
        doGet(request, response);
    }

    protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
        try {
            print(response);
        } catch (DocumentException e) {
            e.printStackTrace();
        }
    }

    public void print(HttpServletResponse response) throws IOException, DocumentException {
        PdfPTable table = initCmnPdfPTable("a");
        PdfPCell cell;
        BaseFont bfChinese = BaseFont.createFont(BASE_FONT, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
        Font headFont2 = new Font(bfChinese, 12, Font.NORMAL);// 设置字体大小
        List<String> lines = new ArrayList<String>();
        lines.add("   hello    ");
        lines.add("");
        lines.add("");
        lines.add("");
        lines.add("           world     ");
        lines.add("");
        lines.add("                ");
        lines.add("");
        lines.add("              你好         ");
        lines.add("            ");
        lines.add("");
        lines.add("                     中国伟大的祖国                  ");
        lines.add("");
        lines.add("本证实书仅对存款人开户证实，不得作为质押权利凭证");
        lines.add("");
        for (String row : lines) {
            cell = new PdfPCell(new Paragraph(row, headFont2));
            cell.setBorder(0);
            cell.setFixedHeight(18);
            cell.setHorizontalAlignment(Element.ALIGN_LEFT);
            table.addCell(cell);
        }
        printPdfTable(table, response);
    }

    public PdfPTable initCmnPdfPTable(String title) throws IOException, DocumentException {
        PdfPTable table = new PdfPTable(new float[]{1000f});// 建立一个pdf表格
        table.setSpacingBefore(160f);// 设置表格上面空白宽度
        table.setTotalWidth(1000);// 设置表格的宽度
        table.setLockedWidth(false);// 设置表格的宽度固定
        table.getDefaultCell().setBorder(0);//设置表格默认为无边框
        BaseFont bfChinese = BaseFont.createFont(BASE_FONT, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
        Font headFont1 = new Font(bfChinese, 14, Font.BOLD);// 设置字体大小
        PdfPCell blankCell = new PdfPCell(new Paragraph());
        blankCell.setBorder(0);
        blankCell.setFixedHeight(20);
        table.addCell(blankCell);
        PdfPCell cell = new PdfPCell(new Paragraph(title, headFont1));
        cell.setBorder(0);
        cell.setFixedHeight(15);//单元格高度
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);// 设置内容水平居中显示
        cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
        table.addCell(cell);
        blankCell.setFixedHeight(5);
        table.addCell(blankCell);
        return table;
    }

    // 单页打印
    public void printPdfTable(PdfPTable table, HttpServletResponse response) throws DocumentException, IOException {
        Document document = new Document(PageSize.A4, 16, 16, 36, 90);
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        PdfWriter writer = PdfWriter.getInstance(document, bos);
        writer.setPageEvent(new PdfPageEventHelper());
        document.open();
        document.add(table);
        document.close();
        response.reset();
        ServletOutputStream out = response.getOutputStream();
        response.setContentType("application/pdf");
//        response.setHeader("Content-disposition", "inline"); //直接在浏览器中打开
        response.setHeader("Content-disposition", "attachment; filename=" + System.currentTimeMillis() + ".pdf");
        response.setContentLength(bos.size());
        response.setHeader("Cache-Control", "max-age=30");
        bos.writeTo(out);
        out.flush();
        out.close();
    }
}