package Controller;

import Model.CompilaRapporti;
import Model.GetDatiFromMattinale;
import Model.SaveStatisticsMailConverted;
import javafx.beans.value.ChangeListener;
import javafx.beans.value.ObservableValue;
import javafx.concurrent.Task;
import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.ProgressBar;
import javafx.scene.control.Tooltip;
import javafx.scene.image.ImageView;
import javafx.scene.input.*;
import javafx.scene.text.Text;
import sample.Main;

import java.awt.*;
import java.io.File;
import java.io.IOException;
import java.net.URL;
import java.text.DecimalFormat;
import java.util.ResourceBundle;

public class MainController implements Initializable {

    private static String FileInserito;
    private static String CartellaDelFileInserito;
    private static String CartellaDeiRapportiCreatiXLS;
    private static String CartellaDeiRapportiCreatiPDF;
    private static String nomRap;
    private static int contatore;
    private static int extTime;
    private static boolean processoFinito = false;
    private static boolean dropInCorso;
    Main main = new Main();




    @FXML
    private ImageView ImmagineXlsVuoto;

    @FXML
    private Button avvioButton;

    @FXML
    private Text testo;

    @FXML
    private ImageView ImmagineXlsPiena;


    @FXML
    private Text testoRapporto;

    @FXML
    private Text tempoStimato;

    @FXML
    private Text elementiCreati;

    @FXML
    private ProgressBar progressBar;

    @FXML
    private Text tempoEffett;

    @FXML
    private Text percentualeText;

    private int inizio;
    private int fine;




    @FXML
    void handleDragDropped(DragEvent event) {
        boolean b = false;
        System.out.println("prima "+b);
        Dragboard db = event.getDragboard();
        File file = db.getFiles().get(0);
        b = true;

        setFileInserito(file.getAbsolutePath());
        setCartellaDelFileInserito(file.getParent());

        dropInCorso = false;


        if (b) {
            fileAccettato();

            System.out.println(b);
            testo.setText(getFileInserito());

            ImmagineXlsVuoto.setVisible(false);
            ImmagineXlsPiena.setVisible(true);



        }
        System.out.println("Il percorso del file caricato "+file.getAbsolutePath());
        System.out.println("La sua cartella: "+file.getParent());






    }

    @FXML
    void showInfo(MouseEvent event) throws IOException {
        main.ShowInfoView();

    }

    void fileAccettato()
    {

        Tooltip.install(avvioButton, new Tooltip("Avvia il processo"));

        ImmagineXlsVuoto.setVisible(false);
        ImmagineXlsPiena.setVisible(true);



        avvioButton.setDisable(false);

    }


    @FXML
    void handleDragOver2(DragEvent event) {




        ImmagineXlsVuoto.setVisible(false);
        ImmagineXlsPiena.setVisible(true);

    }



    @FXML
    void handleDragOver(DragEvent event) {


        if(event.getDragboard().hasFiles()) {
            event.acceptTransferModes(TransferMode.ANY);



        }



    }

    @FXML
    void handleDragOverNOT(DragEvent event) {

        if(dropInCorso)
        {

            ImmagineXlsVuoto.setVisible(true);
            ImmagineXlsPiena.setVisible(false);
        } else
        {
            ImmagineXlsVuoto.setVisible(false);
            ImmagineXlsPiena.setVisible(true);
        }



    }

    @Override
    public void initialize(URL url, ResourceBundle resourceBundle) {
        dropInCorso = true;
        percentualeText.setVisible(false);
        processoFinito=false;

    }











    @FXML
    void avviaProcesso(MouseEvent event) throws Exception {
        avvioButton.setDisable(true);
        setInizio(0);
        long startTime = System.currentTimeMillis();
        percentualeText.setVisible(true);
        processoFinito = false;



        System.out.println("sono stato cliccato");

        //creo la cartella "Cartella Rapporti" nello stesso percorso in cui è il mattinale
        new File(getCartellaDelFileInserito()+"\\Cartella Rapporti Excel").mkdir();
        setCartellaDeiRapportiCreatiXLS(getCartellaDelFileInserito()+"\\Cartella Rapporti Excel");

        //creo la cartella "Cartella Rapporti" nello stesso percorso in cui è il mattinale
        new File(getCartellaDelFileInserito()+"\\Cartella Rapporti PDF").mkdir();
        setCartellaDeiRapportiCreatiPDF(getCartellaDelFileInserito()+"\\Cartella Rapporti PDF");
        Desktop.getDesktop().open(new File(getCartellaDelFileInserito()+"\\Cartella Rapporti PDF\\"));

        GetDatiFromMattinale gDm = new GetDatiFromMattinale();


        try {
            gDm.metodoCheFaTutto();
        } catch (Exception e) {
            e.printStackTrace();
            Alert al = new Alert(Alert.AlertType.ERROR);
            al.setContentText("Errore nel caricamento del File");
            al.show();
        }


        CompilaRapporti compRap = new CompilaRapporti();
        compRap.metodoCheGeneraIFile();
        tempoStimato.setText(compRap.getTempostimato());

        SaveStatisticsMailConverted savStat = new SaveStatisticsMailConverted();
        savStat.recuperaFoglioStatistiche();
        savStat.compilaExcelStatistiche(compRap.getListaNomiFile().size());




        Task<Void> task = new Task<Void>() {
            @Override
            protected Void call() throws Exception {





                for (int i=0; i<compRap.getListaNomiFile().size(); i++) {
                    setInizio(i);

                    System.out.println("1 - KKKKKKKKKKKKKKKKK "+getInizio());

                    //updateMessage(listOfFile[i].getName());
                   /* Thread t1 = new Thread(new Runnable() {

                        @Override
                        public void run() {
                            String s = null;
                            s = compRap.getListaNomiFile().get(getInizio());

                            try {
                                compRap.metodoCheCompilaIRapporti();
                            } catch (Exception e) {
                                e.printStackTrace();
                            }

                            // setNomRap(s);

                            System.out.println((getInizio()+1)+".a) "+s);

                        }
                    });
                    //  System.out.println((getInizio()+1)+".b) "+i);

                    // setNomRap(compRap.getListaNomiFile().get(i));

                    t1.start();*/

                    if(getInizio()>0)
                    {
                        compRap.metodoCheCompilaIRapporti(getInizio());
                        System.out.println("fatto");
                    }
                    System.out.println("2 - KKKKKKKKKKKKKKKKK "+getInizio());

                    System.out.println("LISTA NOMI FILE -->"+compRap.getListaNomiFile().get(i)+ " <---");
                    setFine(compRap.getListaNomiFile().size());
                    updateProgress(i+1, compRap.getListaNomiFile().size());
                    System.out.println("3 - KKKKKKKKKKKKKKKKK "+getInizio());

                    System.out.println((getInizio()+1)+".c) "+i);
                    System.out.println("4 - KKKKKKKKKKKKKKKKK "+getInizio());


                    updateMessage(compRap.getListaNomiFile().get(i));
                    System.out.println("5 - KKKKKKKKKKKKKKKKK "+getInizio());
                    Thread.sleep(3300);
                    System.out.println("6 - KKKKKKKKKKKKKKKKK "+getInizio());

                    if (getInizio()==(compRap.getListaNomiFile().size())-1)
                    {

                        System.out.println("LO SETTOOOOOOO "+processoFinito+" perchè getInizio = "+getInizio()+" e lunghezzaLista = "+(compRap.getListaNomiFile().size()-1));
                        processoFinito = true;
                    }

                    if(processoFinito)
                    {
                        avvioButton.setDisable(false);
                        updateMessage("PROCESSO COMPLETATO");

                        long finishTime = (System.currentTimeMillis())-startTime;

                        long millis = finishTime % 1000;
                        long second = (finishTime / 1000) % 60;
                        long minute = (finishTime / (1000 * 60)) % 60;
                        long hour = (finishTime / (1000 * 60 * 60)) % 24;



                        tempoEffett.setText(String.format("%02d:%02d:%02d.%d", hour, minute, second, millis));

                        processoFinito=false;



                    }

                    System.out.println("il processo, all'interazione '"+ i +"' è finito ? "+processoFinito);

                }
                return null;
            }
        };
        System.out.println("Contatore GetInizio : "+getInizio());

        compRap.metodoCheCompilaIRapporti(getInizio());
        //   Thread.sleep(3300);

        System.out.println("7 - KKKKKKKKKKKKKKKKK "+getInizio());

        task.messageProperty().addListener(new ChangeListener<String>() {
            @Override
            public void changed(ObservableValue<? extends String> observableValue, String oldSt, String newSt) {

                testoRapporto.setText(newSt+" ...");

            }
        });

        System.out.println("8 - KKKKKKKKKKKKKKKKK "+getInizio());

        progressBar.progressProperty().unbind();
        progressBar.progressProperty().bind(task.progressProperty());

        Thread th = new Thread(task);
        th.setDaemon(true);
        th.start();

        //  compRap.salvaUltimoIdRapport();
        compRap.chiudiTuttiFile();

        // compRap.eliminaFileExcel();


        /********************************************************************************/
        Task<Void> task2 = new Task<Void>() {
            @Override
            protected Void call() throws Exception {


                for (int i=0; i<compRap.getListaNomiFile().size(); i++) {
                    setInizio(i);

                    float ciclo = (float)i+1;
                    float dim = (float)compRap.getListaNomiFile().size();

                    DecimalFormat df = new DecimalFormat("#.##");


                    updateMessage(String.valueOf(df.format((ciclo/dim)*100)));

                    Thread.sleep(3300);



                }
                return null;
            }
        };


        task2.messageProperty().addListener(new ChangeListener<String>() {
            @Override
            public void changed(ObservableValue<? extends String> observableValue, String oldSt, String newSt2) {

                percentualeText.setText(newSt2+"%");

            }
        });

        Thread th2 = new Thread(task2);
        th2.setDaemon(true);
        th2.start();

        /********************************************************************************/
        Task<Void> task3 = new Task<Void>() {
            @Override
            protected Void call() throws Exception {


                for (int i=0; i<compRap.getListaNomiFile().size(); i++) {


                    updateMessage(String.valueOf(i+1));

                    Thread.sleep(3300);



                }
                return null;
            }
        };


        task3.messageProperty().addListener(new ChangeListener<String>() {
            @Override
            public void changed(ObservableValue<? extends String> observableValue, String oldSt, String newSt2) {

                elementiCreati.setVisible(true);
                elementiCreati.setText(newSt2+"/"+compRap.getListaNomiFile().size());

            }
        });

        Thread th3 = new Thread(task3);
        th3.setDaemon(true);
        th3.start();


    }






    public String aggiustaTimer(long a)
    {
        if (a<10)
        {
            String g = "0"+a;
            return g;
        }

        return String.valueOf(a);
    }

    public static String getFileInserito() {
        return FileInserito;
    }

    public static void setFileInserito(String fileInserito) {
        FileInserito = fileInserito;
    }

    public static String getCartellaDelFileInserito() {
        return CartellaDelFileInserito;
    }

    public static void setCartellaDelFileInserito(String cartellaDelFileInserito) {
        CartellaDelFileInserito = cartellaDelFileInserito;
    }

    public static String getCartellaDeiRapportiCreatiXLS() {
        return CartellaDeiRapportiCreatiXLS;
    }

    public static void setCartellaDeiRapportiCreatiXLS(String cartellaDeiRapportiCreatiXLS) {
        CartellaDeiRapportiCreatiXLS = cartellaDeiRapportiCreatiXLS;
    }

    public static String getNomRap() {
        return nomRap;
    }

    public static void setNomRap(String nomRap) {
        MainController.nomRap = nomRap;
    }

    public static int getContatore() {
        return contatore;
    }

    public static void setContatore(int contatore) {
        MainController.contatore = contatore;
    }

    public static int getExtTime() {
        return extTime;
    }

    public static void setExtTime(int extTime) {
        MainController.extTime = extTime;
    }

    public static String getCartellaDeiRapportiCreatiPDF() {
        return CartellaDeiRapportiCreatiPDF;
    }

    public static void setCartellaDeiRapportiCreatiPDF(String cartellaDeiRapportiCreatiPDF) {
        CartellaDeiRapportiCreatiPDF = cartellaDeiRapportiCreatiPDF;
    }

    public Text getTestoRapporto() {
        return testoRapporto;
    }

    public int getInizio() {
        return inizio;
    }

    public void setInizio(int inizio) {
        this.inizio = inizio;
    }

    public int getFine() {
        return fine;
    }

    public void setFine(int fine) {
        this.fine = fine;
    }


}
