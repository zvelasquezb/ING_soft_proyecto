package excel;

import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;

public class SCP_controller {

    static String matriz1[][];
    static boolean mat[][];
    int x1 = 0, y1 = 0;
    String url = "";

    public SCP_controller(String asd) {
        url = asd;
    }

    void iniMatriz(int X1, int Y1) {
        x1 = X1;
        y1 = Y1;
        matriz1 = new String[X1][Y1];
        mat = new boolean[X1][Y1];

    }

    void connect() {
        try {
            Document document = null;
            try {
                document = Jsoup.connect(url).get();
            } catch (IOException ex) {
                Logger.getLogger(SCP_controller.class.getName()).log(Level.SEVERE, null, ex);
            }

            for (Element row : document.select("table.tablesorter.full tr")) {
                if (row.select("td:nth-of-type(1)").text().equals("")) {
                    continue;
                } else {
                    String ticker
                            = row.select("td:nth-of-type(1)").toString();

                    String name
                            = row.select("td:nth-of-type(2)").text();
                    String tempPrice
                            = row.select("td.right:nth-of-type(3)").text();
                    String tempPrice1
                            = tempPrice.replace(",", "");
//                  final double price = Double.parseDouble(tempPrice1);
                    for (int i = 0; i < x1; i++) {
                        if (ticker.contains(matriz1[i][0])) {
                            mat[i][0] = true;
                            if (row.select("td:nth-of-type(2)").toString().contains(matriz1[i][1])) {
                                //System.out.println("ok");
                                mat[i][1] = true;
                            } else {
                                matriz1[i][1] = name;
                                mat[i][1] = false;
                                //System.out.println("no");
                            }
                            if (row.select("td.right:nth-of-type(3)").text().contains(matriz1[i][2])) {
                                //System.out.println("ok");
                                mat[i][2] = true;
                            } else {
                                mat[i][2] = false;
                                matriz1[i][2] = tempPrice1;
                                //System.out.println("no");
                            }

                            //System.out.println(ticker+" "+matriz1[0][0]+" "+"ok");
                        }
                    }
                }
            }
        } catch (Exception ex) {
            ex.printStackTrace();
        }

    }

    void loadMatriz(String a[][]) {
        for (int i = 0; i < x1; i++) {

            for (int j = 0; j < y1; j++) {
                matriz1[i][j] = a[i][j];
            }
        }
    }

    void printMatriz1() {
        for (int i = 0; i < matriz1.length; i++) {
            for (int j = 0; j < matriz1[0].length; j++) {
                System.out.print(matriz1[i][j] + " ");
            }
            System.out.println("");
        }
    }

    void printMat() {
        for (int i = 0; i < mat.length; i++) {
            for (int j = 0; j < mat[0].length; j++) {
                System.out.print(mat[i][j] + " ");
            }
            System.out.println("");
        }
    }
}
