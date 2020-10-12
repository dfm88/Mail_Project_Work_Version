package Model;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;

public class XlsToPDF {
    private String percorsoMacro;
    private String percorsoOrigineExcel;
    private String percorsoDestinazionePDF;



    public void writeNewScript(String percorsoOrigineExcel, String percorsoDestinazionePDF) throws IOException
    {

        File macro = new File("");
     //   setPercorsoMacro(macro.getAbsolutePath()+"\\src\\ExcelToPDFConverter\\macro.vbs"); // C:-Users-user-Documents-PROFaxMailProj++\ExcelToPFDConverter\macro.vbs

        setPercorsoMacro("X:\\Sezione Radio\\Operating.Sez.Radio\\Margoni\\Margoni Personale\\Master\\FaxMailProjMASTER\\macro.vbs");

        macro = new File(getPercorsoMacro());


        FileWriter fw = new FileWriter(macro, false);

        BufferedWriter bw = new BufferedWriter(fw);



        bw.write(compilaScript(percorsoOrigineExcel, percorsoDestinazionePDF));
        bw.close();

        try {
            Runtime.getRuntime().exec( "wscript \""+getPercorsoMacro()+"\"" );
        }
        catch( IOException e ) {
            System.out.println(e);

        }

    }

    public String compilaScript(String fielXLS, String filePdf)
    {
        return 	"Dim Excel\r\n" +
                "Dim ExcelDoc\r\n" +
                "\r\n" +
                "Set Excel = CreateObject(\"Excel.Application\")\r\n" +
                "\r\n" +
                "'Open the Document\r\n" +
                "Set ExcelDoc = Excel.Workbooks.open(\""+fielXLS+"\")\r\n" +
                "Excel.ActiveSheet.ExportAsFixedFormat 0, \""+filePdf+"\" ,0, 1, 0,,,0\r\n" +
                "Excel.ActiveWorkbook.Close\r\n" +
                "Excel.Application.Quit";
    }

    public String getPercorsoMacro() {
        return percorsoMacro;
    }

    public void setPercorsoMacro(String percorsoMacro) {
        this.percorsoMacro = percorsoMacro;
    }

    public String getPercorsoOrigineExcel() {
        return percorsoOrigineExcel;
    }

    public void setPercorsoOrigineExcel(String percorsoOrigineExcel) {
        this.percorsoOrigineExcel = percorsoOrigineExcel;
    }

    public String getPercorsoDestinazionePDF() {
        return percorsoDestinazionePDF;
    }

    public void setPercorsoDestinazionePDF(String percorsoDestinazionePDF) {
        this.percorsoDestinazionePDF = percorsoDestinazionePDF;
    }
}
