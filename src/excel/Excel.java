package excel;

import static excel.Excel_controller.compararArchivosExcel;
import static excel.Excel_controller.crearArchivoExcel;
import static excel.Excel_controller.imp;
import static excel.Excel_controller.leerArchivoExcel;
import java.io.File;
import java.io.IOException;


public class Excel {
    
    public static void main(String[] args) throws IOException {
        File f =new File("C:/Users/Zamir/Desktop/Nuevo Hoja de cálculo de Microsoft Excel.xlsx");
        if (f.exists()) {
            excel2_controller obj =new excel2_controller(f);
            Excel_controller obj2= new Excel_controller();
            //System.out.println(obj.x+" "+obj.y);
            //obj.printMatriz1();
            SCP_controller scp =new SCP_controller("https://web.archive.org/web/20190104110157/http://shares.telegraph.co.uk/indices/?index=MCX");
            scp.iniMatriz(obj.x, obj.y);
            scp.loadMatriz(obj.matriz1);
            //System.out.println(scp.matriz1[0][1]);
            scp.connect();
           // System.out.println(scp.matriz1[0][0]+" "+scp.matriz1[0][1]+" "+scp.matriz1[0][2]);
            
            //obj2.modificarArchivoExcel("copia", 0, 3, 2,"hola");
            obj.printMatriz1();
            System.out.println("-------------------------");
            scp.printMatriz1();
            System.out.println("-------------------------");
            scp.printMat();
            
            
           // for (int i = 0; i < scp.mat.length; i++) {
            //    for (int j = 0; j < scp.mat[0].length; j++) {
                    if (scp.mat[1][2]==false) {
                        
                    obj2.modificarArchivoExcel("copia", 0, scp.matriz1, scp.mat);    
                        
                    }
             //   }
            //}
            
        }
    }
    

}
