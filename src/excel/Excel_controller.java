
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
    public static String matriz1[][];
    public static String matriz2[][];
    public static int col1=0,fil1=0;
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
            XSSFSheet sheet = wb.getSheetAt(1);
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

    public static void modificarArchivoExcel(String nombreDelArchivo,int hoja,String comp[][],boolean mat[][]) throws IOException {
            
        try {
            FileInputStream files = new FileInputStream(new File("C:\\Users\\Zamir\\Desktop\\Nuevo Hoja de cálculo de Microsoft Excel.xlsx"));
            
            XSSFWorkbook wb = new XSSFWorkbook(files);
            XSSFSheet sheet = wb.getSheetAt(hoja);
            for (int i = 0; i < mat.length; i++) {
                for (int j = 0; j < mat[0].length; j++) {
                    if (!mat[i][j]) {
            XSSFRow fila= sheet.getRow(i);
            if (fila==null) {
                fila= sheet.createRow(i);
            }
            XSSFCell celda=fila.createCell(j);
            if (celda==null) {
                celda=fila.createCell(j);
            }
            celda.setCellValue(comp[i][j]); 
                    }
                }
            }
            
            files.close();
            FileOutputStream out = new FileOutputStream(nombreDelArchivo + ".xlsx");
            wb.write(out);
            out.close();
            
        } catch (IOException ex) {
            Logger.getLogger(Excel.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    
    public static void compararArchivosExcel(String ruta1,String ruta2) throws IOException {
        try {
            
            FileInputStream files = new FileInputStream(new File(ruta1));
            
            XSSFWorkbook wb = new XSSFWorkbook(files);
            XSSFSheet sheet = wb.getSheetAt(0);
            
            int numFilas=sheet.getLastRowNum();
            //v b
            Row filas= sheet.getRow(1);
                int numColumnss =filas.getLastCellNum();
            matriz1 = new String[sheet.getLastRowNum()][numColumnss];
            //
            for (int i = 0; i < numFilas; i++) {
                Row fila= sheet.getRow(i);
                int numColumns =fila.getLastCellNum();
                 
                for (int j = 0; j < numColumns; j++) {
                    Cell celda = fila.getCell(j);
                    switch (celda.getCellTypeEnum().toString()){
                        case "NUMERIC":
                            //System.out.print(celda.getNumericCellValue()+" ");
                            matriz1[i][j]=String.valueOf(celda.getNumericCellValue());
                            break;
                            
                        case "STRING":
                            //System.out.print(celda.getStringCellValue()+" ");
                            matriz1[i][j]=celda.getStringCellValue();
                            break;
                        
                        case "FORMULA":
                            //System.out.print(celda.getCellFormula()+" ");
                            matriz1[i][j]=celda.getCellFormula();
                            break;
                    }
                    
                }
                //System.out.println("");
            }
        } catch (IOException ex) {
            Logger.getLogger(Excel.class.getName()).log(Level.SEVERE, null, ex);
        
        }
        //
        try {
            
            FileInputStream files = new FileInputStream(new File(ruta2));
            
            XSSFWorkbook wb = new XSSFWorkbook(files);
            XSSFSheet sheet = wb.getSheetAt(0);
            
            int numFilas=sheet.getLastRowNum();
           
            Row filas= sheet.getRow(1);
                int numColumnss =filas.getLastCellNum();
            matriz2 = new String[sheet.getLastRowNum()][numColumnss];
            //
            for (int i = 0; i < numFilas; i++) {
                Row fila= sheet.getRow(i);
                int numColumns =fila.getLastCellNum();
                 
                for (int j = 0; j < numColumns; j++) {
                    Cell celda = fila.getCell(j);
                    switch (celda.getCellTypeEnum().toString()){
                        case "NUMERIC":
                            //System.out.print(celda.getNumericCellValue()+" ");
                            matriz2[i][j]=String.valueOf(celda.getNumericCellValue());
                            break;
                            
                        case "STRING":
                            //System.out.print(celda.getStringCellValue()+" ");
                            matriz2[i][j]=celda.getStringCellValue();
                            break;
                        
                        case "FORMULA":
                            //System.out.print(celda.getCellFormula()+" ");
                            matriz2[i][j]=celda.getCellFormula();
                            break;
                    }
                    
                }
                //System.out.println("");
            }
        } catch (IOException ex) {
            Logger.getLogger(Excel.class.getName()).log(Level.SEVERE, null, ex);
        
        }
        //
        
    }
    
    public static void imp(){
    for (int i = 0; i < matriz1[1].length; i++) {
        for (int j = 0; j < matriz1.length; j++) {
            //System.out.println(matriz1[j][i]);
           // if (matriz1[j][i]==matriz2[j][i]) {
                System.out.println(matriz1[j][i]);
            //}
        }
    }
}
    
    public static void loadMatriz1(String ruta1) throws IOException{
        try {
            
            FileInputStream files = new FileInputStream(new File(ruta1));
            
            XSSFWorkbook wb = new XSSFWorkbook(files);
            XSSFSheet sheet = wb.getSheetAt(0);
            
            int numFilas=sheet.getLastRowNum();
            
            Row filas= sheet.getRow(1);
                int numColumnss =filas.getLastCellNum();
                
            matriz1 = new String[sheet.getLastRowNum()][numColumnss];
            matriz2 = new String[sheet.getLastRowNum()][numColumnss];
            //System.out.println(numFilas+" "+numColumnss);
            fil1=numColumnss;
            //
            for (int i = 0; i < numFilas; i++) {
                Row fila= sheet.getRow(i);
                int numColumns =fila.getLastCellNum();
                 
                for (int j = 0; j < numColumns; j++) {
                    Cell celda = fila.getCell(j);
                    switch (celda.getCellTypeEnum().toString()){
                        case "NUMERIC":
                            //System.out.print(celda.getNumericCellValue()+" ");
                            matriz1[i][j]=String.valueOf(celda.getNumericCellValue());
                            break;
                            
                        case "STRING":
                            //System.out.print(celda.getStringCellValue()+" ");
                            matriz1[i][j]=celda.getStringCellValue();
                            break;
                        
                        case "FORMULA":
                            //System.out.print(celda.getCellFormula()+" ");
                            matriz1[i][j]=celda.getCellFormula();
                            break;
                    }
                    
                }
                //System.out.println("");
            }
        } catch (IOException ex) {
            Logger.getLogger(Excel.class.getName()).log(Level.SEVERE, null, ex);
        
        }
    }
    
    public static void loadMatriz2ex(String a[][]){
        for (int i = 0; i < fil1; i++) {
            matriz1[i][0]=a[i][0];
        }
    }
}
