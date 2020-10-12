package Model;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;


public class FileLoader {

    private FileInputStream fis;
    private Workbook wb;
    private Sheet sh;
    private FileOutputStream fos;

    public FileLoader()
    {

    }



    public FileInputStream getFileInputStream(String percorso) throws IOException {

        return new FileInputStream(percorso);

    }

    public Workbook getExcelWorkBook(FileInputStream Fis) throws IOException {

        return WorkbookFactory.create(Fis);

    }

    public Sheet getExcelSheet(Workbook Wb) throws IOException {

        String nomePagina = Wb.getSheetName(0); //prende il nome della pagina excel

        return Wb.getSheet(nomePagina); //nome della pagina di excel da considerare

    }






}
