
package Model;

        import org.apache.poi.ss.usermodel.*;

        import java.io.FileInputStream;
        import java.io.FileOutputStream;
        import java.io.IOException;
        import java.net.InetAddress;
        import java.net.UnknownHostException;
        import java.sql.Timestamp;
        import java.util.Date;

public class SaveStatisticsMailConverted {

    private static FileLoader fl = new FileLoader();
    private static Workbook wb;
    private static Sheet sh;
    private static FileOutputStream fos;
    private static FileInputStream fis;


    public void recuperaFoglioStatistiche() throws IOException {
        fis = fl.getFileInputStream("X:\\Sezione Radio\\Operating.Sez.Radio\\Margoni\\Margoni Personale\\Master\\FaxMailProjMASTER\\Statistiche\\Statistics.xlsx");
        wb = fl.getExcelWorkBook(fis);
        sh = fl.getExcelSheet(wb);
    }

    public void compilaExcelStatistiche(int numeroRapportiCreati) throws IOException {
        int ultiamRiga = sh.getLastRowNum();
        System.out.println(ultiamRiga);

        Timestamp timestamp = new Timestamp(System.currentTimeMillis());

        Date d = new Date(timestamp.getTime());

        System.out.println(d);


        //dove 5 è il numero di colonne del file statistic
        //DATA - ORA - IP ADDRESS - HOSTNAME - NR RAPPORTI

      /*  sh.createRow(ultiamRiga+1);
        Cell cellaData = sh.getRow(ultiamRiga).getCell(0);
        Cell cellaOra = sh.getRow(ultiamRiga).getCell(1);
        Cell cellaIP = sh.getRow(ultiamRiga).getCell(2);
        Cell cellaHOST = sh.getRow(ultiamRiga).getCell(3);
        Cell cellaNrRapporti = sh.getRow(ultiamRiga).getCell(4);


        cellaData.setCellType.
        cellaOra.setCellFormula("ADESSO()");
        cellaIP.setCellValue(getLocalIP());
        cellaHOST.setCellValue(getLocalHost());
        cellaNrRapporti.setCellValue(numeroRapportiCreati);*/





        sh.createRow(ultiamRiga+1).createCell(0).setCellValue(timestamp.toString().substring(0, 10));
        sh.getRow(ultiamRiga+1).createCell(1).setCellValue(timestamp.toString().substring(11, 19));
        sh.getRow(ultiamRiga+1).createCell(2).setCellValue(getLocalIP());
        sh.getRow(ultiamRiga+1).createCell(3).setCellValue(getLocalHost());
        sh.getRow(ultiamRiga+1).createCell(4).setCellValue(numeroRapportiCreati);

        fos = new FileOutputStream("X:\\Sezione Radio\\Operating.Sez.Radio\\Margoni\\Margoni Personale\\Master\\FaxMailProjMASTER\\Statistiche\\Statistics.xlsx");
        wb.write(fos);

        //evitare che chieda di salvare il file alla chiusura
        wb.getCreationHelper().createFormulaEvaluator().evaluateAll();

        fos.flush();

        fos.close();
        fis.close();

    }



    public String getLocalIP() throws UnknownHostException {
        InetAddress inetAddress = InetAddress.getLocalHost();
        return inetAddress.getHostAddress();

    }

    public String getLocalHost() throws UnknownHostException {
        InetAddress inetAddress = InetAddress.getLocalHost();
        return inetAddress.getHostName();

    }




    public static void main(String[] args) throws IOException {
        SaveStatisticsMailConverted ss = new SaveStatisticsMailConverted();
        ss.recuperaFoglioStatistiche();
        ss.compilaExcelStatistiche(5);
    }

}
