
package excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Excel_controller {

    public Excel_controller() {
    }
    
    
    
        public static void crearArchivoExcel(String nombreDelArchivo, String nombreDelLibro, int Extencion/*if(1){.xls}else if(2){.xlsx}*/) {
        
        if (Extencion == 1) {
            Workbook book = new HSSFWorkbook();//se crea un archivo excel
            Sheet sheet = book.createSheet(nombreDelLibro);
            try {
                FileOutputStream fileout = new FileOutputStream(nombreDelArchivo + ".xls");
                book.write(fileout);
                fileout.close();
                
            } catch (FileNotFoundException ex) {
                Logger.getLogger(Excel.class.getName()).log(Level.SEVERE, null, ex);
            } catch (IOException ex) {
                Logger.getLogger(Excel.class.getName()).log(Level.SEVERE, null, ex);
            }            
        } else if (Extencion == 2) {
            Workbook book = new XSSFWorkbook();//se crea un archivo excel
            Sheet sheet = book.createSheet(nombreDelLibro);
            try {
                FileOutputStream fileout = new FileOutputStream(nombreDelArchivo + ".xlsx");
                book.write(fileout);
                fileout.close();
                
            } catch (FileNotFoundException ex) {
                Logger.getLogger(Excel.class.getName()).log(Level.SEVERE, null, ex);
            } catch (IOException ex) {
                Logger.getLogger(Excel.class.getName()).log(Level.SEVERE, null, ex);
            }            
        }
        
    }
    
    public static void leerArchivoExcel(String ruta) throws IOException {
        try {
            FileInputStream files = new FileInputStream(new File(ruta));
            
            XSSFWorkbook wb = new XSSFWorkbook(files);
            XSSFSheet sheet = wb.getSheetAt(0);
            
            int numFilas=sheet.getLastRowNum();
            for (int i = 0; i < numFilas; i++) {
                Row fila= sheet.getRow(i);
                int numColumns =fila.getLastCellNum();
                for (int j = 0; j < numColumns; j++) {
                    Cell celda = fila.getCell(j);
                    switch (celda.getCellTypeEnum().toString()){
                        case "NUMERIC":
                            System.out.print(celda.getNumericCellValue()+" ");
                            break;
                            
                        case "STRING":
                            System.out.print(celda.getStringCellValue()+" ");
                            break;
                        
                        case "FORMULA":
                            System.out.print(celda.getCellFormula()+" ");
                            break;
                    }
                    
                }
                System.out.println("");
            }
        } catch (IOException ex) {
            Logger.getLogger(Excel.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    public static void modificarArchivoExcel(String nombreDelArchivo,int hoja,int row,int column,String modificacion) throws IOException {
            
        try {
            FileInputStream files = new FileInputStream(new File("C:\\Users\\Zamir\\Desktop\\Notas de clase\\Ingenieria de software\\Trabajos\\parcial1\\Lextura excel\\Nueva carpeta\\productos.xlsx"));
            
            XSSFWorkbook wb = new XSSFWorkbook(files);
            XSSFSheet sheet = wb.getSheetAt(hoja);
            XSSFRow fila= sheet.getRow(row);
            if (fila==null) {
                fila= sheet.createRow(row);
            }
            XSSFCell celda=fila.createCell(column);
            if (celda==null) {
                celda=fila.createCell(column);
            }
            celda.setCellValue(modificacion); 
            files.close();
            FileOutputStream out = new FileOutputStream(nombreDelArchivo + ".xlsx");
            wb.write(out);
            out.close();
            
        } catch (IOException ex) {
            Logger.getLogger(Excel.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
}
