package excel;

import static excel.Excel_controller.crearArchivoExcel;
import java.io.IOException;


public class Excel {
    
    public static void main(String[] args) throws IOException {
        Excel_controller excel= new Excel_controller();
        
        crearArchivoExcel("P", "hola", 2);
        //leerArchivoExcel("C:\\Users\\Zamir\\Desktop\\Notas de clase\\Ingenieria de software\\Trabajos\\parcial1\\Lextura excel\\Nueva carpeta\\productos.xlsx");
    }
    

}
