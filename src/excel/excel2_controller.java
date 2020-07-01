
package excel;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class excel2_controller {
int x,y;
String matriz1 [][];
static String matriz2 [][];
    public excel2_controller(File filename) {
        List cellData=new ArrayList(); 
    try{
        FileInputStream fileimputstram=new FileInputStream(filename);
        
        XSSFWorkbook workBook= new XSSFWorkbook  (fileimputstram);     
        
        XSSFSheet hssfSheet = workBook.getSheetAt(0);
        
        Iterator rowIterator =hssfSheet.rowIterator();
        
        while(rowIterator.hasNext()){
        
            XSSFRow hssfRow =(XSSFRow) rowIterator.next();
            Iterator iterator=hssfRow.cellIterator();
            List cellTemp=new ArrayList();
            while(iterator.hasNext()){
            XSSFCell hssCell=(XSSFCell) iterator.next();
            cellTemp.add(hssCell);
            }
            cellData.add(cellTemp);
                    }
        
    }catch (Exception e){
        e.printStackTrace();
    }
    obtener(cellData);
    fillMatriz1(cellData);
    }
    void obtener(List cellDataList){
        for (int i = 0; i < cellDataList.size(); i++) {
            x=cellDataList.size();
            List celTemplist   =(List) cellDataList.get(i);
            for (int j = 0; j < celTemplist.size(); j++) {
                y=celTemplist.size();
                XSSFCell hssfCell  =(XSSFCell) celTemplist.get(j);
                String stringCellValue= hssfCell.toString();
                
               // System.out.print(stringCellValue+" ");
            }
            //System.out.println();
        }
        matriz1=new String [x][y];
        matriz2=new String [x][y];
    }
    void fillMatriz1(List cellDataList){
        for (int i = 0; i < cellDataList.size(); i++) {
            x=cellDataList.size();
            List celTemplist   =(List) cellDataList.get(i);
            for (int j = 0; j < celTemplist.size(); j++) {
                y=celTemplist.size();
                XSSFCell hssfCell  =(XSSFCell) celTemplist.get(j);
                String stringCellValue= hssfCell.toString();
                matriz1[i][j]=stringCellValue;
                //System.out.print(stringCellValue+" ");
            }
            //System.out.println();
        }
    }
    void printMatriz1(){
        for (int i = 0; i < matriz1.length; i++) {
            for (int j = 0; j < matriz1[0].length; j++) {
                System.out.print(matriz1[i][j]+" ");
            }
            System.out.println("");    
        }
    }
public static void loadMatriz2ex(String a[][]){
        for (int i = 0; i < matriz2.length; i++) {
            for (int j = 0; j < matriz2[0].length; j++) {
            matriz2[i][j]=a[i][j];    
            }
        }
    }
}
