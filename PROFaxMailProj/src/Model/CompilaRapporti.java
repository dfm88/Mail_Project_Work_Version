package Model;

import Controller.MainController;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.sql.Timestamp;
import java.util.ArrayList;

public class CompilaRapporti
{
    private static Workbook wb;
    private static Sheet sh;
    private static FileInputStream fis;
    private static FileOutputStream fos;
    private static Row row;
    private static Cell cell;
    private static Timestamp timestamp = new Timestamp(System.currentTimeMillis());
    private static GetDatiFromMattinale getDatiFromMattinale= new GetDatiFromMattinale();
    private static ArrayList<String> listaNomiFile;
    private int IDRapporto;
    private String percorsoMaster;
    final String estensioneExcel = ".xlsx";
    final String estensionePDF = ".pdf";
    private static FileLoader fll;
    private static String tempostimato;






    public void metodoCheGeneraIFile() throws Exception {

        fll = new FileLoader();

        //getDatiFromMattinale.metodoCheFaTutto();
        File f = new File("");
        System.out.println("Absolute path "+f.getAbsolutePath());
        setPercorsoMaster("X:\\Sezione Radio\\Operating.Sez.Radio\\Margoni\\Margoni Personale\\Master\\FaxMailProjMASTER\\Util_Excel_Files\\MASTERRapportiVigilanza.xlsx");

        System.out.println(getPercorsoMaster());
        //apertura del file RAPPORTIVIGILANZA
        fis = fll.getFileInputStream(getPercorsoMaster());
        wb = fll.getExcelWorkBook(fis);
        sh = fll.getExcelSheet(wb);

        //Recupero l'ultimo IDRapporto Salvato 80761
        int id = (int) sh.getRow(11).getCell(5).getNumericCellValue();
        setIDRapporto(id);

        listaNomiFile = new ArrayList<String>();

        for (int i = 0; i<getDatiFromMattinale.getListaRagioniSociali().size(); i++) {
            //	sh = wb.cloneSheet(0);
            String nomePag = wb.getSheetName(0); //nome della pagina di excel da considerare

            sh = wb.getSheet(nomePag);


            File masterFile = new File(MainController.getFileInserito());

            String ragioneSoc = getDatiFromMattinale.getListaRagioniSociali().get(i);
            String orario = getDatiFromMattinale.getListaDateEOra().get(i).toLocaleString();


            String nomeFile = ragioneSoc;


            nomeFile = nomeFile.substring(0, nomeFile.indexOf("\n"));


            nomeFile = nomeFile.replaceAll("\\:", " ");
            nomeFile = nomeFile.replaceAll("\\?", " ");
            nomeFile = nomeFile.replaceAll("\\^", " ");
            nomeFile = nomeFile.replaceAll("\\\\", " ");
            nomeFile = nomeFile.replaceAll("\"", " ");
            nomeFile = nomeFile.replaceAll("\\/", " ");
            nomeFile = nomeFile.replaceAll("\\*", " ");
            nomeFile = nomeFile.replaceAll("\\’", "'");
            nomeFile = nomeFile.replaceAll("\\–", "-");
            nomeFile = nomeFile.replaceAll("\\|", " ");
            nomeFile = nomeFile.replaceAll("\n", " ").replace("\r", "");
            orario = orario.replaceAll("\\:", " ");

            String nomePagineOrario = nomeFile + " " + orario;


            //   File nuovoFile = new File(getPercorsoMaster()+"\\WHATAFUCK\\"+nomePagineOrario+".xlsx");


            listaNomiFile.add(nomePagineOrario + estensioneExcel);

            System.out.println("aa " + masterFile.toPath());
            //  System.out.println("bb "+nuovoFile.toPath());
            // System.out.println("bb "+nuovoFile.toPath());

            // if(!nuovoFile.exists())

            //  Files.copy(masterFile.toPath(), nuovoFile.toPath());  ***A COSA SERVE? BOOO

            fis.close();

            fos = new FileOutputStream(getPercorsoMaster());
            wb.write(fos);

            fos.close();
        }

        System.out.println("MetodoCheGeneraIlFile ha finito");

        setTempostimato(tempoStimato(listaNomiFile.size()));





    }

    public String tempoStimato(int quantitàRapporti)
    {
        long tempoStim = quantitàRapporti*3300;

        long millis = tempoStim % 1000;
        long second = (tempoStim / 1000) % 60;
        long minute = (tempoStim / (1000 * 60)) % 60;
        long hour = (tempoStim / (1000 * 60 * 60)) % 24;

        return String.format("%02d:%02d:%02d.%d", hour, minute, second, millis);
    }

    public String avanzamentoRapporti(int indice, int massimo)
    {
        return indice+" / "+massimo;
    }




    public void metodoCheCompilaIRapporti(int i) throws Exception
    {

        //    for(int i = 0; i< getDatiFromMattinale.getListaRagioniSociali().size();i++) {
        //salvo e incremento l'IDRapporto
        fis = new FileInputStream(getPercorsoMaster());
        wb = WorkbookFactory.create(fis);

        //Recupero l'ultimo IDRapporto Salvato 80761
        sh = wb.getSheet(wb.getSheetName(0));




        Cell cellIDRapporto = sh.getRow(11).getCell(5);
        setIDRapporto(getIDRapporto() + 1);
        cellIDRapporto.setCellValue(getIDRapporto());


        //salvo la ragione sociale
        Cell cellRagioneSociale = sh.getRow(4).getCell(8);

        cellRagioneSociale.setCellValue(getDatiFromMattinale.getListaRagioniSociali().get(i));


        //salvo data e ora dell'evento
        Cell cellDataEOra = sh.getRow(17).getCell(4);

        cellDataEOra.setCellValue(getDatiFromMattinale.getListaDateEOra().get(i).toLocaleString());

        //faccio il time stamp della creazione del file
        Cell cellTimeStamp = sh.getRow(8).getCell(1);

        cellTimeStamp.setCellValue(timestamp.toString().substring(0, 19));//per tagliare i millesimi

        //salvo il tipo di segnalazione
        Cell cellTipoSegnalazione = sh.getRow(18).getCell(2);

        cellTipoSegnalazione.setCellValue(getDatiFromMattinale.getListaTipoAllarme().get(i));


        //salvo l'esito
        Cell cellEsito = sh.getRow(22).getCell(2);

        cellEsito.setCellValue(getDatiFromMattinale.getListaEsiti().get(i));

        String percorsoRapportExcelCreato = MainController.getCartellaDeiRapportiCreatiXLS() + "\\" + listaNomiFile.get(i);


        //  System.out.println(cellRagioneSociale.toString());




        fos = new FileOutputStream(percorsoRapportExcelCreato);
        wb.write(fos);

        //evitare che chieda di salvare il file alla chiusura

        fos.flush();


        fos.close();

        fis.close();


        XlsToPDF xls2PDF = new XlsToPDF();
        String nomeFileSenzaXLS = listaNomiFile.get(i).replaceAll(".xlsx", "");
        System.out.println("scrittura script");
        // xls2PDF.writeNewScript(percorsoRapportExcelCreato, getPercorsoMaster()+"\\Pdf_WHATAFUCK\\"+listaNomiFile.get(i));
        //JAVAFX
        xls2PDF.writeNewScript(percorsoRapportExcelCreato, MainController.getCartellaDeiRapportiCreatiPDF() + "\\" + nomeFileSenzaXLS + estensionePDF);
        // Thread.sleep(3300);
        //   }
        salvaUltimoIdRapport();
        wb.getCreationHelper().createFormulaEvaluator().evaluateAll();
        wb.close();
        fos.close();	//********** ATTENZIONE NUOVI PER PROVARE CHIUSURA FILE *************
    }

    public void metodoCheConverte()
    {

    }

    public void salvaUltimoIdRapport() throws IOException {
        //Scrivo l'ultimo ID nel master 80761
        sh = wb.getSheet(wb.getSheetName(0));
        sh.getRow(11).getCell(5).setCellValue(getIDRapporto());
        System.out.println("Ultimo id creato "+getIDRapporto());

        System.out.println("ho generato nr '"+getDatiFromMattinale.getListaEsiti().size()+"' rapporti");
        fos = new FileOutputStream(getPercorsoMaster());
        wb.write(fos);

        //evitare che chieda di salvare il file alla chiusura
        wb.getCreationHelper().createFormulaEvaluator().evaluateAll();

        fos.flush();



        fos.close();
    }

    public void chiudiTuttiFile() throws IOException, IOException {
        //evitare che chieda di salvare il file alla chiusura
        wb.getCreationHelper().createFormulaEvaluator().evaluateAll();
        wb.close();
        fos.close();	//********** ATTENZIONE NUOVI PER PROVARE CHIUSURA FILE *************
        fis.close();
        // System.exit(0);
    }

    public void SCriptMOthaFucka() throws IOException {
        for (int i=0 ; i<getDatiFromMattinale.getListaRagioniSociali().size(); i++)
        {
            XlsToPDF xls2PDF = new XlsToPDF();
            File ff = new File(getPercorsoMaster() + "\\Excel_WHATAFUCK\\" + listaNomiFile.get(i));
            xls2PDF.writeNewScript(ff.getAbsolutePath(), getPercorsoMaster()+"\\Pdf_WHATAFUCK\\"+listaNomiFile.get(i));

        }
    }

    public void eliminaFileExcel()
    {
        for (int i=0 ; i<getDatiFromMattinale.getListaRagioniSociali().size(); i++)
        {

            File ff = new File(getPercorsoMaster() + "\\Excel_WHATAFUCK\\" + listaNomiFile.get(i));
            ff.delete();
        }
        System.exit(0);
    }





    public static void main(String[] args) throws Exception
    {
        getDatiFromMattinale.metodoCheFaTutto();

        CompilaRapporti cmprap = new CompilaRapporti();


        cmprap.metodoCheGeneraIFile();

        // cmprap.metodoCheCompilaIRapporti();




        //  cmprap.metodoCheAttivaLaMacro();

        cmprap.chiudiTuttiFile();

        //  cmprap.SCriptMOthaFucka();

        cmprap.eliminaFileExcel();

    }

    public String getPercorsoMaster() {
        return percorsoMaster;
    }

    public void setPercorsoMaster(String percorsoMaster) {
        this.percorsoMaster = percorsoMaster;
    }

    public static Workbook getWb() {
        return wb;
    }

    public static void setWb(Workbook wb) {
        CompilaRapporti.wb = wb;
    }

    public static Sheet getSh() {
        return sh;
    }

    public static void setSh(Sheet sh) {
        CompilaRapporti.sh = sh;
    }

    public static FileInputStream getFis() {
        return fis;
    }

    public static void setFis(FileInputStream fis) {
        CompilaRapporti.fis = fis;
    }

    public static FileOutputStream getFos() {
        return fos;
    }

    public static void setFos(FileOutputStream fos) {
        CompilaRapporti.fos = fos;
    }

    public static Row getRow() {
        return row;
    }

    public static void setRow(Row row) {
        CompilaRapporti.row = row;
    }

    public static Cell getCell() {
        return cell;
    }

    public static void setCell(Cell cell) {
        CompilaRapporti.cell = cell;
    }

    public static Timestamp getTimestamp() {
        return timestamp;
    }

    public static void setTimestamp(Timestamp timestamp) {
        CompilaRapporti.timestamp = timestamp;
    }

    public static GetDatiFromMattinale getGetDatiFromMattinale() {
        return getDatiFromMattinale;
    }

    public static void setGetDatiFromMattinale(GetDatiFromMattinale getDatiFromMattinale) {
        CompilaRapporti.getDatiFromMattinale = getDatiFromMattinale;
    }

    public static ArrayList<String> getListaNomiFile() {
        return listaNomiFile;
    }

    public static void setListaNomiFile(ArrayList<String> listaNomiFile) {
        CompilaRapporti.listaNomiFile = listaNomiFile;
    }

    public int getIDRapporto() {
        return IDRapporto;
    }

    public void setIDRapporto(int IDRapporto) {
        this.IDRapporto = IDRapporto;
    }


    public static String getTempostimato() {
        return tempostimato;
    }

    public static void setTempostimato(String tempostimato) {
        CompilaRapporti.tempostimato = tempostimato;
    }
}
