package eu.vancl.martin.exceltest1;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Scanner;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;

public class ExcelTest {

    public ExcelTest() {
        CellReference cr;
        Cell prijmeni;
        Cell jmeno;
        Cell rok;
        Row row;
        
        try {
            POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream("vykaz.xls"));
            HSSFWorkbook wb = new HSSFWorkbook(fs);
            HSSFSheet sheet = wb.getSheet("072016"); // jmeno listu

            cr = new CellReference("C9"); // souradnice bunky
            row = sheet.getRow(cr.getRow());
            prijmeni = row.getCell(cr.getCol());

            cr = new CellReference("H9"); // souradnice bunky
            row = sheet.getRow(cr.getRow());
            jmeno = row.getCell(cr.getCol());

            cr = new CellReference("C12"); // souradnice bunky
            row = sheet.getRow(cr.getRow());
            rok = row.getCell(cr.getCol());

            cr = new CellReference("C5"); // souradnice bunky
            row = sheet.getRow(cr.getRow());

            System.out.println("PRIJMENI: " + prijmeni.getStringCellValue());
            System.out.println("JMENO: " + jmeno.getStringCellValue());
            System.out.println("ROK: " + rok.getNumericCellValue());
            sheet = wb.getSheet("072016");

            Scanner s = new Scanner(System.in);
            System.out.println("Nove prijmeni: ");
            String prijmeniNove = s.next();
            System.out.println("Nove jmeno: ");
            String jmenoNove = s.next();
            System.out.println("Novy rok: ");
            int rokNovy = s.nextInt();

            prijmeni.setCellValue(prijmeniNove);
            jmeno.setCellValue(jmenoNove);
            rok.setCellValue(rokNovy);

            FileOutputStream outputStream = new FileOutputStream(new File("vykaz2.xls"));
            wb.write(outputStream);
            outputStream.close();

            //prijmeni.setCellValue("Prijmeni");
            //jmeno.setCellValue("Jmeno");
            //rok.setCellValue(2017);
            
            // nezavirat okno terminalu ve windows
            System.out.println("\n ...neco zmackni...");
            System.in.read(); 
        } catch (Exception e) {
            System.err.println(e);
        }
    }

    public static void main(String[] args) {
        new ExcelTest();
    }
}
